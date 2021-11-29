use calamine::{open_workbook, Xlsx, Reader};
use anyhow::{anyhow, Result};
use clap::Parser;
use tokio::{fs::OpenOptions, io::AsyncWriteExt};
use std::path::Path;
use std::ffi::OsStr;
use std::str::FromStr;

// 以下部分用于处理 CLI

// 定义 excel-img 的 CLI 的主入口，它包含若干个子命令
// 下面 /// 的注释是文档，clap 会将其作为 CLI 的帮助


/// 使用 Rust 原生实现的 excel 批量下载 img 的 cli 程序
#[derive(Parser, Debug)]
#[clap(version = "1.0", author = "Steve Wong <ikook.dev@gmail.com>")]
struct Opts {
    #[clap(subcommand)]
    subcmd: SubCommand,
}

#[derive(Parser, Debug)]
enum SubCommand {
    Save(Save)
}

/// feed get with an file path and we will retrieve the response for you
#[derive(Parser, Debug)]
struct Save {
    /// excel 文件路径
    path: String,

    /// 文件解析列数，格式为："x,y"，x 为文件名列；y 为图片url列
    #[clap(parse(try_from_str = parse_kv_pair))]
    body: KvPair,

    /// 图片保存路径
    save_path: String,
}

#[derive(Debug)]
struct KvPair {
    sort_id: usize,
    url: usize,
}

/// 当我们实现 FromStr trait 后，可以用 str.parse() 方法将字符串解析成 KvPair
impl FromStr for KvPair {
    type Err = anyhow::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        // 使用 = 进行 split，这会得到一个迭代器
        let mut split = s.split(",");
        let err = || anyhow!(format!("Failed to parse {}", s));
        Ok(Self {
            // 从迭代器中取第一个结果作为 key，迭代器返回 Some(T)/None
            // 我们将其转换成 Ok(T)/Err(E)，然后用 ? 处理错误
            sort_id: (split.next().ok_or_else(err)?).to_string().parse::<usize>().unwrap(),
            // 从迭代器中取第二个结果作为 value
            url: (split.next().ok_or_else(err)?).to_string().parse::<usize>().unwrap(),
        })
    }
}

/// 因为我们为 KvPair 实现 FromStr，这里可以直接 s.parse() 得到 KvPair
fn parse_kv_pair(s: &str) -> Result<KvPair> {
    Ok(s.parse()?)
}

/// 处理 path 子命令
async fn path(args: &Save) -> Result<()> {
    let mut excel: Xlsx<_> = open_workbook(&args.path).unwrap();

    if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
        let mut i = 0;
        println!("处理总数为：{:?}", r.get_size().0 - 1);
        for row in r.rows() {
            if i == 0 {
                i = i + 1;
                continue;
            } else {
                println!("正在处理第{}个文件", i);
                i = i + 1;
            }


            let url = row[args.body.url].to_string();

            let mut filename = args.save_path.to_string();
            let sort_id = row[args.body.sort_id].to_string();

            let c = ".".to_string();
            let mut response  = get(&url).await?;
            let suffix_o = get_extension_from_filename(&url);
            let suf = suffix_o.unwrap().to_string();

            filename = filename + &sort_id + &c + &suf;
            save(&filename, &mut response).await?;
        }

        println!("处理完毕！！！");
    }
    Ok(())
}

fn get_extension_from_filename(filename: &str) -> Option<&str> {
    Path::new(filename)
        .extension()
        .and_then(OsStr::to_str)
}

async fn save(filename: &str, response: &mut reqwest::Response) -> Result<()> {
    let mut options = OpenOptions::new();
    let mut file = options
        .append(true)
        .create(true)
        .read(true)
        .open(filename)
        .await?;

    while let Some(chunk) = &response.chunk().await.expect("Failed") {
        match file.write_all(&chunk).await {
            Ok(_) => {}
            Err(e) => return Err(anyhow!("File {} save error: {}", filename, e)),
        }
    }
    Ok(())
}

async fn get(url: &str) -> Result<reqwest::Response> {
    reqwest::get(url)
        .await
        .map_err(|e| anyhow!("Request url {} error: {}", url, e))
}

#[tokio::main]
async fn main() -> Result<()> {
    let opts: Opts = Opts::parse();

    let result = match opts.subcmd {
        SubCommand::Save(ref args) => path(args).await?,
    };

    Ok(result)
}

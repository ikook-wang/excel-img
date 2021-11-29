use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
use anyhow::{anyhow, Result};
use clap::Parser;
use reqwest::{header, Client, Response, Url};
use std::{collections::HashMap, str::FromStr};

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
    path(Path)
}

/// feed get with an file path and we will retrieve the response for you
#[derive(Parser, Debug)]
struct Path {
    /// 需要解析的文件路径
    #[clap(parse(try_from_str = parse_path))]
    path: String,
}

fn parse_path(s: &str) -> Result<String> {
    // 这里我们仅仅检查一下 URL 是否合法
    let _url: Url = s.parse()?;

    Ok(s.into())
}


fn main() {
    println!("Hello, world!");
}

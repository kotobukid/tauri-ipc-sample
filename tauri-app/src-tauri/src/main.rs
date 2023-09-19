// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::collections::HashMap;
use std::process::{Command, Stdio};
use std::io::{Read, Write};
use std::path::Path;
use std::fs::read;
use serde::Serialize;
use tauri::Manager;
use serde_json::to_string;

// Learn more about Tauri commands at https://tauri.app/v1/guides/features/command
#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

struct ExcelCol {
    current: Vec<char>,
}

impl ExcelCol {
    fn new() -> ExcelCol {
        ExcelCol {
            current: vec!['A'],
        }
    }

    fn next(&mut self) -> String {
        let s = self.current.iter().collect::<String>();

        let mut carry = true;
        for ch in self.current.iter_mut().rev() {
            if carry {
                *ch = ((*ch as u8) + 1) as char;
                if *ch > 'Z' {
                    *ch = 'A';
                } else {
                    carry = false;
                }
            }
        }

        if carry {
            self.current.insert(0, 'A');
        }

        s
    }
}

#[derive(Serialize)]
struct WriteExcelInfo<'a> {
    excel_name: &'a str,
    cell: HashMap<String, String>
}

impl WriteExcelInfo<'_> {
    fn new(values: Vec<String>) -> Self {
        let mut excel_col = ExcelCol::new();

        WriteExcelInfo {
            excel_name: "test123",
            cell: values.iter().map(|s| {
                (format!("{}{}", excel_col.next(), 1), s.clone())
            }).collect()
        }
    }
}

#[tauri::command]
async fn write_excel(values: Vec<String>) -> tauri::Result<Vec<u8>> {

    let config: WriteExcelInfo = WriteExcelInfo::new(values);

//     let xlsx_config = r###"
//     {
//     "excel_name":"test123",
//     "cell":
//     {
//         "D8":"D8に入れたい内容",
//         "F8":"F8に入れたい内容"
//     }
// }
//     "###;

    let xlsx_config = to_string(&config).unwrap();
    println!("{}", xlsx_config);

    let mut child = Command::new("node")
        .arg("../../script.js")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("Failed to start child process");

    let mut stdin = child.stdin.take().expect("Failed to open stdin");
    let mut stdout = child.stdout.take().expect("Failed to open stdout");

    stdin.write_all(xlsx_config.as_bytes()).expect("Failed to write to stdin");
    // stdin.write_all(b"Hello from Rust\n").expect("Failed to write to stdin");

    let mut buffer = String::new();
    stdout.read_to_string(&mut buffer).expect("Failed to read from stdout");

    println!("Received from Node.js: {}", buffer);

    let _ = child.wait().expect("Failed to wait on child");
    read_file_binary(buffer).await
}

async fn read_file_binary(path_str: String) -> tauri::Result<Vec<u8>> {
    let path = Path::new(path_str.as_str());
    let data = read(&path)?;
    Ok(data)
}

fn main() {
    tauri::Builder::default()
        .setup(|app| {
            // 開発時だけdevtoolsを表示する。
            #[cfg(debug_assertions)]
            app.get_window("main").unwrap().open_devtools();

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![greet, write_excel])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

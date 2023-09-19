use std::process::{Command, Stdio};
use std::io::{Read, Write};

fn main() {
    let mut child = Command::new("node")
        .arg("../script.js")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("Failed to start child process");

    let mut stdin = child.stdin.take().expect("Failed to open stdin");
    let mut stdout = child.stdout.take().expect("Failed to open stdout");

    stdin.write_all(b"Hello from Rust\n").expect("Failed to write to stdin");

    let mut buffer = String::new();
    stdout.read_to_string(&mut buffer).expect("Failed to read from stdout");

    println!("Received from Node.js: {}", buffer);

    let _ = child.wait().expect("Failed to wait on child");
}

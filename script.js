process.stdin.on('data', function(data) {
    console.log(`Received data from Rust: ${data.toString()}`);
    process.stdout.write("Hello from Node.js\n");

    process.exit(0);
});

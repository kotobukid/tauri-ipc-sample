import fs from 'node:fs';
import process from "node:process";

import XlsxPopulate from "xlsx-populate";

process.stdin.on('data', async (data: string) => {

    const checkAndCreateDir = async (dirname_relative: string): Promise<void> => {
        try {
            const stat = await fs.promises.stat(dirname_relative);
            return
        } catch (error) {
            // @ts-ignore
            if (error.code === 'ENOENT') {
                const created = await fs.promises.mkdir(dirname_relative);
            } else {
                throw error;
            }
        }
    }

    try {
        // console.log(`Received data from Rust: ${data.toString()}`);
        // process.stdout.write("Hello from Node.js\n");

        const json = JSON.parse(data);
        //. テンプレート
        const templateFile = './01.xlsx';

        const dir_created = await checkAndCreateDir("./tmp/");

        //. 出力先(Herokuではtmp以下にファイルを一時保存する)
        const outputFilePath = './tmp/' + json.excelName + '.xlsx';

        const book = await XlsxPopulate.fromFileAsync(templateFile)

        // 0番目のsheetに書き込む
        const sheet1 = book.sheet(0);
        for (const [key, value] of Object.entries(json.cell)) {
            // @ts-ignore
            sheet1.cell(key).value(value);
        }
        const result = await book.toFileAsync(outputFilePath);

        process.stdout.write(outputFilePath);
        // process.stdout.write(result);

        process.exit(0);
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
});

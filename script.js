var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import fs from 'node:fs';
import process from "node:process";
import XlsxPopulate from "xlsx-populate";
process.stdin.on('data', (data) => __awaiter(void 0, void 0, void 0, function* () {
    const checkAndCreateDir = (dirname_relative) => __awaiter(void 0, void 0, void 0, function* () {
        try {
            const stat = yield fs.promises.stat(dirname_relative);
            return;
        }
        catch (error) {
            // @ts-ignore
            if (error.code === 'ENOENT') {
                const created = yield fs.promises.mkdir(dirname_relative);
            }
            else {
                throw error;
            }
        }
    });
    try {
        // console.log(`Received data from Rust: ${data.toString()}`);
        // process.stdout.write("Hello from Node.js\n");
        const json = JSON.parse(data);
        //. テンプレート
        const templateFile = './01.xlsx';
        const dir_created = yield checkAndCreateDir("./tmp/");
        //. 出力先(Herokuではtmp以下にファイルを一時保存する)
        const outputFilePath = './tmp/' + json.excelName + '.xlsx';
        const book = yield XlsxPopulate.fromFileAsync(templateFile);
        // 0番目のsheetに書き込む
        const sheet1 = book.sheet(0);
        for (const [key, value] of Object.entries(json.cell)) {
            // @ts-ignore
            sheet1.cell(key).value(value);
        }
        const result = yield book.toFileAsync(outputFilePath);
        process.stdout.write(outputFilePath);
        // process.stdout.write(result);
        process.exit(0);
    }
    catch (e) {
        console.error(e);
        process.exit(1);
    }
}));

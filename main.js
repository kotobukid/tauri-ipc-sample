const XlsxPopulate = require('xlsx-populate');
const express = require('express');
const app = express();

app.use(express.urlencoded({
    extended: true
}));
app.use(express.json());

app.listen(process.env.PORT || 3000);
console.log('Server is online.');

app.post('/createExcel', function (req, res) {
    console.log(req.body);
    const json = req.body;
    //. テンプレート
    const templateFile = './01.xlsx';
    //. 出力先(Herokuではtmp以下にファイルを一時保存する)
    const outputFilePath = '/tmp/' + json.excelName + '.xlsx';
    XlsxPopulate.fromFileAsync(templateFile).then(book => {
        // 0番目のsheetに書き込む
        var sheet1 = book.sheet(0);
        for (const [key, value] of Object.entries(json.cell)) {
            sheet1.cell(key).value(value);
        }
        // outputFilePathnに書き出し
        book.toFileAsync(outputFilePath).then(result => {
            // resで端末にダウンロードさせる
            res.download(outputFilePath);
        });
    });
})

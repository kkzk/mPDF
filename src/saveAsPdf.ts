const winax = require("winax");
import path = require("path");

import * as fo from './fileOrder';
import * as fs from "fs";

const xlTypePDF = 0;
const xlQualityStandard = 0;
const xlQualityMinimum = 1;
const xlQuality = xlQualityStandard;

export function saveExcel(data: fo.Document[], workspaceDir: string, xlsRelpath: string) {
    const xlsPath = path.resolve(path.join(workspaceDir, xlsRelpath));
    var pdfPath = path.join(workspaceDir, ".mPDF", xlsRelpath);
    var basename = path.basename(pdfPath);
    var name = path.parse(basename).name;
    pdfPath = path.join(
        path.dirname(pdfPath),
        `${name}.pdf`
    );
    console.log("convert from ", xlsPath, " to ", pdfPath);

    const excel = new winax.Object("Excel.Application", { activate: true});
    const workbooks = excel.Workbooks;
    const workbook = workbooks.Open(xlsPath, 0, true);
    const doc = data.find((doc) => doc.name === xlsRelpath);
    const sheets = doc?.worksheets?.filter((sheetname) => sheetname.visible);
//    console.log(sheets);

    var replace = true;
    sheets?.forEach((e) => {
        workbook.sheets[e.name].Select(replace);
        replace = false;
    });
    fs.mkdirSync(path.dirname(pdfPath), { recursive: true });
    workbook.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdfPath, xlQuality);
    workbook.Saved = true;
    workbook.Close();
    excel.Quit();
}

export function savePdf(jsonData: fo.Document[], workspaceDir: string, documentUri: string) {
    const document = decodeURI(documentUri);
    const ext = path.extname(document);
    console.log(`extension: ${ext}`);
    switch(ext) {
        case ".xlsx":
            saveExcel(jsonData, workspaceDir ,document);
            break;
        default:
            break;
    }
}

import * as XLSX from 'xlsx'
declare let window: any;
declare let FileReader: any;
declare let document: any;
var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";

export const isBrower = 'undefined' !== typeof window;
/**
 * 读出JSON为
 * @param file 
 * @param success 
 */
export function readAsJSON(file: string | any, success?: (d: any) => void): any {
    let Data: any = {};
    if (isBrower) {
        var reader = new FileReader();
        reader.onload = function (e: any) {
            var data = e.target.result;
            let workbook = XLSX.read(data, { type: 'binary' })
            workbook.SheetNames.forEach((d: string) => {
                Data[d] = XLSX.utils.sheet_to_json(workbook.Sheets[d])
            })
            if (success instanceof Function) {
                success(Data)
            }
        }
        if (rABS) reader.readAsBinaryString(file);
        else reader.readAsArrayBuffer(file);
    } else {
        let workbook = XLSX.readFile(file)
        workbook.SheetNames.forEach((d: string) => {
            Data[d] = XLSX.utils.sheet_to_json(workbook.Sheets[d])
        })
        if (success instanceof Function) {
            success(Data)
        }
    }
    return Data;
}
/**
 * 
 * @param Data 
 * @param FileName 
 */
export function writeFileFromJSON(Data: any, FileName: string) {
    let WorkBook = XLSX.utils.book_new()
    for (let x in Data) {
        XLSX.utils.book_append_sheet(WorkBook, XLSX.utils.json_to_sheet(Data[x]))
    }
    if (isBrower) {
        XLSX.writeFile(WorkBook, FileName)
    } else {
        XLSX.writeFile(WorkBook, FileName)
    }
}
/**
 * 
 * @param id 
 * @param FileName 
 */
export function writeFileFromTable(id: string, FileName: string) {
    if (isBrower)
        XLSX.writeFile(XLSX.utils.table_to_book(document.getElementById(id)), FileName)
    else
        throw new Error('ONLY SUPPORT BROWER')
}
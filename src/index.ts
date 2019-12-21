import * as XLSX from 'xlsx'
declare let window: any;
declare let FileReader: any;
declare let document: any;
var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";

export const isBrower = 'undefined' !== typeof window;
/**
 * 触发弹出文件选择框
 * @param accept 
 */
export function select_file(accept = "*") {
    return new Promise((s, j) => {
        let i = document.createElement('input');
        i.type = 'file';
        i.accept = accept;
        i.hidden = true;
        document.body.appendChild(i);
        i.onchange = (ev) => {
            if (i.files && i.files.length > 0) {
                s(i.files);
            } else {
                j('NoFile')
            }
            document.body.removeChild(i);
        }
        i.click();
    })
}
export class KeyMap {
    default: string | boolean | number = "";
    code: string = "";
    name: string = "";
}
/**
 * 读取Excel文件为JSON内容
 * @param {string} file 可选参数，传入File对象或者留空则在在前端自动弹出文件选择框 
 * @param {KeyMap} map 
 */
export function readAsJSON(file?: string | any, map?: KeyMap[]): Promise<{ [index: string]: { [index: string]: string | number | "" } } | { [index: string]: string | number | boolean }[]> {
    return new Promise(async (s, j) => {
        let Data: any = {};
        if (isBrower) {
            if (!file) {
                file = (await select_file('.xlsx,.xlx'))[0]
            }
            var reader = new FileReader();
            reader.onload = function (e: any) {
                var data = e.target.result;
                let workbook = XLSX.read(data, { type: 'binary' })
                if (workbook.SheetNames.length == 0) {
                    s({});
                }
                workbook.SheetNames.forEach((d: string) => {
                    Data[d] = XLSX.utils.sheet_to_json(workbook.Sheets[d])
                })
                if (map) {
                    let td: { [index: string]: string | number | boolean }[] = [];
                    for (let x of Data[workbook.SheetNames[0]]) {
                        let p = {};
                        for (let o of map) {
                            p[o.code] = x[o.name] || x[o.code] || (x.default instanceof Function ? x.default(x) : x.default)
                        }
                        td.push(p);
                    }
                    s(td);
                } else {
                    s(Data)
                }
            }
            if (rABS) reader.readAsBinaryString(file);
            else reader.readAsArrayBuffer(file);
        } else {
            let workbook = XLSX.readFile(file)
            workbook.SheetNames.forEach((d: string) => {
                Data[d] = XLSX.utils.sheet_to_json(workbook.Sheets[d])
            })
            s(Data)
        }
    })
}
/**
 * 从JSON结构中生成XLSX文件
 * @param Data 
 * @param FileName 
 */
export function writeFileFromJSON(Data: { [index: string]: { [index: string]: string | number | boolean }[] }, FileName: string) {
    let WorkBook = XLSX.utils.book_new()
    for (let x in Data) {
        XLSX.utils.book_append_sheet(WorkBook, XLSX.utils.json_to_sheet(Data[x]))
    }
    XLSX.writeFile(WorkBook, FileName)
}
/**
 * 从Table标签中生成XLSX文件内容
 * @param id 
 * @param FileName 
 */
export function writeFileFromTable(id: string, FileName: string) {
    if (isBrower)
        XLSX.writeFile(XLSX.utils.table_to_book(document.getElementById(id)), FileName)
    else
        throw new Error('ONLY SUPPORT BROWER')
}
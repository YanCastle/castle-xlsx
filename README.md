# castle-xlsx
```shell
npm i -S castle-xlsx
```
# 读取文件 read xlsx file in brower
```html
<input type="file" name="" id="" @change="handleChange">
```
```typescript
import { readAsJSON } from "castle-xlsx";
```
```typescript
handleChange(e: any) {
    const files = e.target.files;
    if (!files) {
        return;
    }
    this.Loading = true;
    readAsJSON(files[0], (d: any) => {
        if (d["Sheet1"] instanceof Array) {
            this.XlsxData = d;
        } else {
            error("找不到Sheet1的表，请确认模版是否正确");
        }
        this.Loading = false;
    });
}
```

# 导出文件 export xlsx file in brower 

## 从table中导出数据，数据源为table内容 export from table
```typescript
    writeFileFronTable('table标签的id值, the table element`s id','文件名称 filename.xlsx')
```
## 从JSON数据导出为xlsx文件 export from json 
```typescript
    wirteFileFromJSON({
        SheetName:[
            [1,2,3,4,5,6],
            [1,2,3,4,5,6],
            [1,2,3,4,5,6],
        ]
    },'文件名称 filename.xlsx')
```
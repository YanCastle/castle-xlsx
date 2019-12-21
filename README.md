# @ctsy/xlsx
```shell
yarn add @ctsy/xlsx
```
# 读取文件 read xlsx file in brower
```html
<!-- 现在可以不用手动写上传控件了，可以直接调用readAsJSON -->
<!-- <input type="file" name="" id="" @change="handleChange"> -->
<button @click="upload">导入(import)</button>
```
```typescript
import { readAsJSON } from "@ctsy/xlsx";
```
```typescript
async upload(e: any) {
    try{
        this.Loading = true;
        let rs = await readAsJSON(undefined,[{name:"标题",code:'键',default:'默认值'}]);
        // rs就是获得的结果
    }catch(e){

    }finally{

        this.Loading = false;
    }
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
            {Title:1,Value:2},
            {Title:1,Value:2},
            {Title:1,Value:2},
            {Title:1,Value:2},
            {Title:1,Value:2},
            {Title:1,Value:2}
        ]
    },'文件名称 filename.xlsx')
```

# castle-xlsx
xlsx
```html
<input type="file" name="" id="" @change="handleChange">
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
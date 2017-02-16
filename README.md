## 更新日志

+ 2017/2/15
 - 支持填入文本文档

## 样例
+ row : number类型,行数量
+ column : number类型, 列数量
+ defaultList :array类型,内部填充json，让单元格变为一个单选框模式
    - x number类型,单选框的x轴坐标
    - y number类型,单选框的x轴坐标
    - value array类型, 数组项为单选框的值
+ value : json类型 key为`x轴-y轴`,value为填入值
+ filePath : string类型,生成文档的保存路径,默认路径为path.join(__dirname,"1.xlsx");
+ cb : function类型,有两个参数,(err,result);

```
//所有的参数都是从1开始，而非从0开始
const  generateXlsx =require('./index.js');
generateXlsx({
    row:3,
    column:3,
    defaultList:[{
        x:2,
        y:1,
        value:["男","女"]
    }],
    value:{
        "1-3":"我的中国"
    },
    filePath:path.join(__dirname,"1.xlsx"),
    cb
});
```
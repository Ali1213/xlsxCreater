
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
    }
});
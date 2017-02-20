var generateOffice = require("./generateOffice");
var path = require("path");
const fs = require("fs");
const constant = require("./constant");





const generateXlsx = function({
    row=0,
    column=0,
    defaultList = [],
    value = {},
    filepath = path.join(__dirname,"1.xlsx")
},
    cb = function(){}){
    let count = 0;
    let contentStringArr = [];
    const office={
        templateDir:path.join(__dirname,'template')
    };
    //一个诡异的数字转英文 0->A,26->AA;
    const toAlphabet = function(num){
        return String.fromCharCode(65 + parseInt(num));
    };
    const to26 = function(num){
        return num / 26 >= 1 ? toAlphabet(num/26-1) + to26(num%26) : toAlphabet(num%26);
    };
    //生成sheet1.xml;
    let s = constant.sheet1;
    const cell = function(row,column){
        let str = '';
        for(let i=0;i<row;i++){
            str += `<row r="${i+1}" spans="1:${column}" x14ac:dyDescent="0.15">`;
            for(let j=0;j<column;j++){
                if(!value[i+1+'-'+(j+1)]){
                    str+= `<c r="${to26(j)+(i+1)}" s="2"></c>`;
                }else{
                    contentStringArr[count] = value[i+1+'-'+(j+1)];
                    str+= `<c r="${to26(j)+(i+1)}" s="2" t="s"><v>${count++}</v></c>`;
                }
            }
            str += "</row>";
        }
        return str;
    };

    if(row && column){
        s = s.replace(/<sheetData>[\s\S]*?<\/sheetData>/i,()=>{

            return `<sheetData>${cell(row,column)}<\/sheetData>`;
        })


        s= s.replace(/<dimension ref=[\s\S]+?\/>/,()=>{
            return `<dimension ref="A1:${to26(row-1)+column}"/>`;
        })
    }

    if(defaultList.length){
        s = s.replace(/<dataValidations[\s\S]*?<\/dataValidations>/i,()=>{
            let sStr = `<dataValidations count="${defaultList.length}">`;
            for(let elem of defaultList){
                sStr += `
                        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${to26(elem.x-1)+elem.y}">
                            <formula1>"${elem.value.join(',')}"</formula1>
                        </dataValidation>
                    `
            }
            return sStr + "</dataValidations>";
        })
    }

    //生成sharedStrings.xml
    let ss = constant.sharedStrings;

    const shared = function(){
        var str = "";
        for (let i=0;i<count;i++){
            str +=`<si>
            <t>${contentStringArr[i]}</t>
            <phoneticPr fontId="1" type="noConversion"/>
                </si>`
        }
        return str;
    };

    if(Object.keys(value)){
        ss = ss.replace(/<sst[\s\S]+?<\/sst>/i,()=>{
            let str = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">`;
            str+=shared();
            return str + "</sst>";
        })
    }


    office.sheet1 = s;
    office.sharedStrings = ss;

    generateOffice(office,filepath,cb)
};



module.exports = generateXlsx;
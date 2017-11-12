const XLSX = require('xlsx'), request = require('request-promise');
const Promise = require('bluebird');
const fs = require('fs');
const GSUtils = require('./GSUtils');

// Return Promise
function parseXLSX (url) {
    return requestServ(url).then((data) => {
        let workbook= XLSX.read(data, {type: 'buffer', cellNF: true, cellStyles: true, cellDates: true});
    let spreadsheets={ properties: {title:url.replace(/^.+\/(\w+)\.\w+$/,"$1") }, sheets:[]};
    for(let ind=0;ind<workbook.SheetNames.length;ind++)
    {
        let sheet ={properties: {
            sheetId:ind,
            title:workbook.SheetNames[ind],
            index:ind }
        };
        let workbooksheet = workbook.Sheets[workbook.SheetNames[ind]];
        if( workbooksheet['!ref']) {
            if(workbooksheet['!merges']) {
                sheet.merges = setMerges(workbooksheet['!merges'],ind);
            }
            let range = separete(workbooksheet['!ref'], ":");
            let propX = GSUtils.alphaToNum(range.ind[0]), propY = parseInt(range.ind[1]);
            sheet.properties.gridProperties = getGridProperties(range.data, propX, propY);
            sheet.data = createGridData(workbooksheet, propX, propY);
        }
        spreadsheets.sheets.push(sheet);
    }
    return spreadsheets;
});
}

function requestServ(url) {
    return request(url, {encoding: null}, function(err, res, data) {
        if (err || res.statusCode !== 200 || res.headers["content-type"] !== "application/vnd.ms-excel"
            || res.headers["content-type"] !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" )  {

            throw err!==null ?  err : " Error dowload files "  ;
        }
        return data;
    });
}

function createGridData(workbooksheet,propX,propY) {
    let gridDada={ startRow:0 ,startColumn:0 };
    gridDada.rowData ={values:[]};
    let csv = XLSX.utils.sheet_to_csv(workbooksheet,{}).split("\n").map((res) => {return res.split(",")});
    csv.forEach((row,indRow) => {
        let arr=[];
    row.forEach((col,indCol) => {
        let cellData={};
    if (col==="") {
        cellData.UserEnteredValue=setExtendedValue(col,"s");
    } else {
        let point = GSUtils.numToAlpha(indCol+propX-1)+(indRow+propY);
        let sheet =workbooksheet[point];
        cellData.UserEnteredValue=setExtendedValue(sheet.v,sheet.t);
        if (sheet.f) {
            cellData.UserEnteredValue=setExtendedValue('='+sheet.f, 'f');
        }
    }
    arr.push(cellData);
});
    gridDada.rowData.values.push(arr)
});
    return gridDada;
}

function setExtendedValue(val,extended) {
    let extendedValue ={};
    switch (extended) {
        case 'n':extendedValue.numberValue=val;
            break;
        case 's':extendedValue.stringValue=val;
            break;
        case 'b':extendedValue.boolValue=val;
            break;
        case 'f':extendedValue.formulaValue=val;
            break;
        default:extendedValue.stringValue=val;
            break;
    }
    return extendedValue;
}
function setMerges(merges,insexSheet) {
    let arr =[];
    merges.forEach((val) => {
        let gridRange ={};
    gridRange.sheetId = insexSheet;
    gridRange.startRowIndex = val.s.r;
    gridRange.endRowIndex = val.e.r;
    gridRange.startColumnIndex = val.s.c;
    gridRange.endColumnIndex = val.e.c;
    arr.push(gridRange);
});
    return arr;
}

function getGridProperties(val,propX,propY) {
    const offset = val.match(/^([A-Z]+)(\d+)$/).slice(1);
    return { columnCount:GSUtils.alphaToNum(offset[0])-propX+1 ,rowCount:parseInt(offset[1])-propY+1 }
}

function separete(val,separ) {
    let [key, value] = val.split(separ);
    return {ind:key.match(/^([A-Z]+)(\d+)$/).slice(1) ,data:value};
}

module.export = parseXLSX;



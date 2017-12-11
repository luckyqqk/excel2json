var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var moment = require('moment');
var glob = require('glob');
var async = require('async');
var config = require('../config.json');

module.exports = {
    /**
     * 文件转json
     * @param fileName {string} 全路径
     * @param outDirArray {string} 导出路径
     */
    toJson: function (fileName, outDirArray) {
        _toJson(xlsx.parse(fileName), outDirArray);
    }
};


/**
 * export .xlsx file to json format.
 * excel: json string converted by 'node-xlsx'。
 * outDirArray : directory for exported json files.
 */
function _toJson(excel, outDirArray) {
    var outJson = {};
    var attrArray = [];
    // 收集页签内数据,最终导出json文件的数据,直接放入outJson.属性修正数据放入attrArray.
    for (var sheetIdx = 0; sheetIdx < excel.worksheets.length; sheetIdx++) {
        var sheet = excel.worksheets[sheetIdx];
        var output = _parseSheet(sheet);
        if (!output)
            continue;
        var sheetName = String(sheet.name);
        var sheetNameArray = sheetName.split(".");
        var fileName = sheetNameArray[0].split("-")[0];
        var attrName = sheetNameArray[1];
        if (!attrName) {    // 本签 : 收集本签信息,考虑到相同本签分页的情况
            if (!outJson[fileName]) {
                outJson[fileName] = output;
            } else {
                if (Array.isArray(outJson[fileName])) {
                    outJson[fileName] = outJson[fileName].concat(output);
                } else {
                    for (let k in output) {
                        outJson[fileName][k] = output[k];
                    }
                }
            }
        } else {    // 属签 : 直接收集属签信息,等所有本签信息收集完毕后,再填充属签内容
            output['fileName'] = fileName;
            output['attrName'] = attrName;
            attrArray.push(output);
        }
    }
    // 将收集到的属性放入对应的数据中, 定位数据( 线索outJson[fileName][key] )后,给该数据增加attrName属性
    attrArray.forEach(attr=> {
        var fileN = attr['fileName'];
        if (!outJson[fileN]) {
            console.warn(`warn: can't find file ${fileN}`);
            return;
        }
        var attrN = attr['attrName'];
        delete attr['fileName'];
        delete attr['attrName'];
        for (var k in attr) {
            if (!outJson[fileN][k]) {
                console.warn(`warn: ${fileN}.${attrN} can't find key ${k}`);
                continue;
            }
            outJson[fileN][k][attrN] = attr[k];
        }
    });
    // fs的代理回调
    var cbAgent = function (fileName) {
        return function (err) {
            if (err) {
                console.error("error：", err);
                throw err;
            }
            console.info(`exported successfully  --> ${fileName}`);
        }
    };

    var writeAgent = function (fileName, toJson) {
        return function (_cb) {
            fs.writeFile(fileName, toJson, _cb);
        };
    };

    // 写入json文件
    for (var key in outJson) {
        var toJson = JSON.stringify(outJson[key], null, 2);
        var writeArray = [];
        outDirArray.forEach(dir=> {
            var fileName = path.resolve(dir, key + ".json");
            writeArray.push(new writeAgent(fileName, toJson));
        });
        async.parallel(writeArray, cbAgent(key + ".json"));
    }

}

var OUTPUT_ARRAY = 0;
var OUTPUT_OBJ_VALUE = 1;
var OUTPUT_OBJ_ARRAY = 2;

/**
 * parse one sheet and return the result as a json object or array
 *
 * @param sheet
 * @returns {*}
 * @private
 */
function _parseSheet(sheet) {
    if (!sheet.data || sheet.data.length < 1)
        return;
    var typeCfg = config.xlsx.type || 2;  // 默认第二行为数据类型行
    var headCfg = config.xlsx.head || 3;  // 默认第三行为头信息行
    var rowType = sheet.data[typeCfg - 1];
    var rowHead = sheet.data[headCfg - 1];

    var outType = OUTPUT_ARRAY;
    var idType = !!rowType && !!rowType[0] ? String(rowType[0].value).trim() : '[]';
    if (idType === 'id') {
        outType = OUTPUT_OBJ_VALUE;
    } else if (idType === 'id[]') {
        outType = OUTPUT_OBJ_ARRAY;
    }
    var output = outType ? {} : [];
    var outRow = 0;   // 供打印导出多少条数据用
    for (var rowIdx = headCfg; rowIdx < sheet.maxRow; rowIdx++) {
        outRow = rowIdx;
        var row = sheet.data[rowIdx];
        // 第一个字段为主键ID,跳过没有主键的行
        if (!row || !row[0] || !String(row[0].value).trim()) {
            //console.log(`${sheet.name}导出结束,共${rowIdx}行`);
            continue;
        }
        var jsonObj = {};
        var id;

        for (var colIdx = 0; colIdx < sheet.maxCol; colIdx++) {
            // 遇到没有头的单元格,就跳过
            if (!rowHead[colIdx] || !rowHead[colIdx].value) continue;  // 判断是否为空
            var colName = String(rowHead[colIdx].value).trim();     // 判断trim后是否为空
            if (!colName) continue;

            var type = !!rowType[colIdx] && !!rowType[colIdx].value ? String(rowType[colIdx].value).toLowerCase().trim() : 'string';

            var cellValue = !!row[colIdx] ? row[colIdx].value : null;
            if (cellValue == null) {
                cellValue = type == 'number' ? 0 : '';
            }
            //var cellValue = !!row[colIdx] && !!String(row[colIdx].value).trim() ? String(row[colIdx].value).trim() : '';
            switch (type) {
                case 'id': // id
                    id = cellValue;
                    break;
                case 'id[]': // id[]
                    id = cellValue;
                    if (!output[id]) {
                        output[id] = [];
                    }
                    break;
                case 'string':
                    cellValue = String(cellValue).trim() || '';
                    if (isDateType(cellValue)) {
                        parseDateType(jsonObj, colName, cellValue);
                    } else {
                        jsonObj[colName] = cellValue || "";
                    }
                    break;
                //case 'date':
                //  parseDateType(jsonObj, colName, cellValue);
                //  break;
                case 'number':
                    jsonObj[colName] = Number(cellValue) || 0;
                    break;
                case 'bool':
                    jsonObj[colName] = Boolean(cellValue) || false;
                    break;
                case '{}': //support {number boolean string date} property type
                    parseObject(jsonObj, colName, cellValue);
                    break;
                case '[]': //[number] [boolean] [string]
                    parseBasicArrayField(jsonObj, colName, cellValue);
                    break;
                case '[{}]':
                    parseObjectArrayField(jsonObj, colName, cellValue);
                    break;
                case 'json':
                    cellValue = cellValue || "";
                    jsonObj[colName] = JSON.parse(cellValue);
                    break;
                default:
                    console.log('********************************************' + type);
                    console.log('unrecognized type by sheet:', sheet.name, cellValue, typeof (cellValue));
                    break;
            }
            //console.log('********************************************');
            //console.log("--->",rowIdx,type,jsonObj[colName], colName,cell);
        }
        if (outType === OUTPUT_OBJ_VALUE) {
            output[id] = jsonObj;
        } else if (outType === OUTPUT_OBJ_ARRAY) {
            output[id].push(jsonObj);
        } else if (outType === OUTPUT_ARRAY) {
            output.push(jsonObj);
        }
    }

    console.log(`${sheet.name}导出结束,共${outRow + 1 - headCfg}行`);
    //console.log("output******",output);
    return output;
}

/**
 * parse date type
 * row:row of xlsx
 * key:col of the row
 * value: cell value
 */
function parseDateType(row, key, value) {
    row[key] = convert2Date(value);
    //console.log(value,key,row[key]);
}

/**
 * convert string to date type
 * value: cell value
 */
function convert2Date(value) {
    var dateTime = moment(value);
    if (dateTime.isValid()) {
        return dateTime.format("YYYY-MM-DD HH:mm:ss");
    } else {
        return "";
    }
}

/**
 * parse object array.
 */
function parseObjectArrayField(row, key, value) {
    if (!value || value.length == 1) {
        row[key] = [];
        return;
    }

    var obj_array = value.split(',');
    //if (value.indexOf(',') !== -1) {
    //  obj_array = value.split(',');
    //} else {
    //  obj_array.push(value.toString());
    //}

    // if (typeof(value) === 'string' && value.indexOf(',') !== -1) {
    //     obj_array = value.split(',');
    // } else {
    //     obj_array.push(value.toString());
    // };

    var result = [];

    obj_array.forEach(function (e) {
        if (!e) return;
        result.push(array2object(e.split(';')));
    });

    row[key] = result;
}

/**
 * parse object from array.
 *  for example : [a:123,b:45] => {'a':123,'b':45}
 */
function array2object(array) {
    var result = {};
    array.forEach(function (e) {
        if (!e) return;
        var kv = e.trim().split(':');
        if (isNumber(kv[1])) {
            kv[1] = Number(kv[1]);
        } else if (isBoolean(kv[1])) {
            kv[1] = toBoolean(kv[1]);
        } else if (isDateType(kv[1])) {
            kv[1] = convert2Date(kv[1]);
        }
        result[kv[0]] = kv[1];
    });
    return result;
}

/**
 * parse object
 */
function parseObject(field, key, data) {
    field[key] = array2object(data.split(';'));
}


/**
 * parse simple array.
 */
function parseBasicArrayField(field, key, array) {
    var basic_array;

    if (typeof array === "string") {
        basic_array = array.length > 1 ? array.split(config.xlsx.arraySeparator || ',') : [];
    } else {
        basic_array = [];
        basic_array.push(array);
    }
    var result = [];
    if (isNumberArray(basic_array)) {
        basic_array.forEach(function (element) {
            result.push(Number(element));
        });
    } else if (isBooleanArray(basic_array)) {
        basic_array.forEach(function (element) {
            result.push(toBoolean(element));
        });
    } else { //string array
        result = basic_array;
    }
    // console.log("basic_array", result + "|||" + cellValue);
    field[key] = result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
    return value.toString().toLowerCase() === 'true';
}

/**
 * is a boolean array.
 */
function isBooleanArray(arr) {
    return arr.every(function (element, index, array) {
        return isBoolean(element);
    });
}

/**
 * is a number array.
 */
function isNumberArray(arr) {
    return arr.every(function (element, index, array) {
        return isNumber(element);
    });
}

/**
 * is a number.
 */
function isNumber(value) {

    if (typeof (value) == "undefined") {
        return false;
    }

    if (typeof value === 'number') {
        return true;
    }
    return !isNaN(+value.toString());
}

/**
 * boolean type check.
 */
function isBoolean(value) {

    if (typeof (value) == "undefined") {
        return false;
    }

    if (typeof value === 'boolean') {
        return true;
    }

    var b = value.toString().trim().toLowerCase();

    return b === 'true' || b === 'false';
}

/**
 * date type check.
 */
function isDateType(value) {
    if (value) {
        var str = value.toString();
        return moment(new Date(value), "YYYY-M-D", true).isValid() || moment(value, "YYYY-M-D H:m:s", true).isValid() || moment(value, "YYYY/M/D H:m:s", true).isValid() || moment(value, "YYYY/M/D", true).isValid();
    }
    return false;
}

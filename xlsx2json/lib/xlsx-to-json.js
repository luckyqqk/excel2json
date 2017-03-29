var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var moment = require('moment');
var glob = require('glob');
var config = require('../config.json');

module.exports = {
  /**
   * 文件转json
   * @param fileName {string} 全路径
   * @param outDir {string} 导出路径
   */
  toJson: function (fileName, outDir) {
    _toJson(xlsx.parse(fileName), outDir);
  }
};


/**
 * export .xlsx file to json format.
 * excel: json string converted by 'node-xlsx'。
 * outDir : directory for exported json files.
 */
function _toJson(excel, outDir) {
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
    var fileName = sheetNameArray[0], attrName = sheetNameArray[1];
    output['fileName'] = fileName;
    if (!attrName) {
      outJson[fileName] = output;
    } else {
      output['attrName'] = attrName;
      attrArray.push(output);
    }
  }
  // 属性修正
  attrArray.forEach((attr)=>{
    var fileName = attr['fileName'];
    var attrName = attr['attrName'];
    var fileJson = outJson[fileName];
    if (!fileJson) {
      console.log(`Error:: attr ${attrName} can't find owner ${fileName}`);
      return;
    }
    delete attr['fileName'];
    delete attr['attrName'];
    for (var k in attr) {
      if (!fileJson[k]) {
        console.log(`Error::file has no data. attr: ${attrName}, key: ${k}`);
        //console.log(JSON.stringify(attr));
        return;
      }
      fileJson[k][attrName] = attr[k];
    }
  });
  // 写入json文件
  for (var key in outJson) {
    var toJson = outJson[key];
    var fileName = toJson['fileName'] + ".json";
    delete toJson['fileName'];
    toJson = JSON.stringify(toJson, null, 2);
    var fileName = path.resolve(outDir, fileName);
    fs.writeFile(fileName, toJson, function (err) {
      if (err) {
        console.log("error：", err);
        throw err;
      }
      console.log('exported successfully  -->  ', path.basename(fileName));
    });
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

  if (!rowType[0] || !rowType[0].value) {
  	console.error(`${sheet.name} has not find type id, toJson fail!`);
  	return;
  }

  var outType = OUTPUT_ARRAY;
  var idType = String(rowType[0].value).trim();
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
    // 第一个字段为主键ID,遇到没有主键的行,视为导出结束
    if (!row || !row[0] || !String(row[0].value).trim()) {
      //console.log(`${sheet.name}导出结束,共${rowIdx}行`);
      break;
    }
    var jsonObj = {};
    var id;

    for (var colIdx = 0; colIdx < sheet.maxCol; colIdx++) {
      // 遇到没有头的单元格,就结束
      if (!rowHead[colIdx] || !rowHead[colIdx].value) break;  // 判断是否为空
      var colName = String(rowHead[colIdx].value).trim();     // 判断trim后是否为空
      if (!colName) break;

      var type = rowType[colIdx];
      type = !type || !type.value ? 'basic' : String(type.value).toLowerCase().trim();
      //console.log(colName,cell.value,typeof(cell),type);
      // if the cell is empty, do not export the key
      //var cell = row[colIdx];
      var cellValue = !!row[colIdx] && !!String(row[colIdx].value).trim() ? String(row[colIdx].value).trim() : '';
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
        case 'basic': // number string boolean date
          if (isDateType(cellValue)) {
            parseDateType(jsonObj, colName, cellValue);
          } else {
            jsonObj[colName] = cellValue || '';
          }
          break;
        case 'date':
          parseDateType(jsonObj, colName, cellValue);
          break;
        case 'string':
          if (isDateType(cellValue)) {
            parseDateType(jsonObj, colName, cellValue);
          } else {
            jsonObj[colName] = cellValue || "";
          }
          break;
        case 'number':
          jsonObj[colName] = Number(cellValue) || 0;
          break;
        case 'bool':
          jsonObj[colName] = Boolean(cellValue) || false;
          break;
        case '{}': //support {number boolean string date} property type
          parseObject(jsonObj, colName, cellValue);
          break;
        case '[]': //[number] [boolean] [string]  todo:support [date] type
          parseBasicArrayField(jsonObj, colName, cellValue);
          break;
        case '[{}]':
          jsonObj[colName] = parseObjectArrayField(jsonObj, colName, cellValue) || [];
          break;
        default:
          console.log('********************************************' + type);
          console.log('unrecognized type', cellValue, typeof (cellValue));
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

  var obj_array = [];

  if (value) {
    if (value.indexOf(',') !== -1) {
      obj_array = value.split(',');
    } else {
      obj_array.push(value.toString());
    }
  }

  // if (typeof(value) === 'string' && value.indexOf(',') !== -1) {
  //     obj_array = value.split(',');
  // } else {
  //     obj_array.push(value.toString());
  // };

  var result = [];

  obj_array.forEach(function (e) {
    if (e) {
      result.push(array2object(e.split(';')));
    }
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
    if (e) {
      var kv = e.trim().split(':');
      if (isNumber(kv[1])) {
        kv[1] = Number(kv[1]);
      } else if (isBoolean(kv[1])) {
        kv[1] = toBoolean(kv[1]);
      } else if (isDateType(kv[1])) {
        kv[1] = convert2Date(kv[1]);
      }
      result[kv[0]] = kv[1];
    }
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
    var arraySeparator = config.xlsx.arraySeparator || ',';
    basic_array = array.split(arraySeparator);
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

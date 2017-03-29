### 作用
让excel表达复杂的json格式,将xlsx文件转成json。

### npm相关
* 如需当做npm模块引用请切换到`npm`分支。

### 感谢
感谢于小懒先生,初始项目down于他的git:https://github.com/koalaylj/xlsx2json
整体代码结构和思路均无变化,更改内容如下:
    1.将表头中的列名和数据类型分离成两行(根据我们的使用习惯,有时候列名过长,若类型掺在其中,影响表的阅读).
    2.重写了读写逻辑,去掉了配表中的页签顺序限制(原来属签一定要排在本签右边).
    3.增加了默认配置和默认数据(没填数据类型的字段(id类型必须填),默认为string类型,默认值为'').


### 使用说明
> 目前只支持.xlsx格式，不支持.xls格式。

* 本项目是基于nodejs的，所以需要先安装nodejs。
* 配置config.json
```json
{
    "xlsx": {
        "type": 2,                              // 数据类型(基本就用id,number,string)
        "head": 3,                              // 表头所在的行，第一行可以是注释，第二行是表头。
        "excelDir": "./excel/**/[^~$]*.xlsx",   // xlsx文件 glob配置风格
        "outDir": "./json",                     //  导出的json存放的位置
        "arraySeparator":","                    // 数组的分隔符
    }
}
```
* 执行`export.sh/export.bat`即可将`./excel/*.xlsx` 文件导成json并存放到 `./json` 下。json名字以excel的sheet名字命名。

* 补充(一般用不上)：
    * 执行`node index.js -h` 查看使用帮助。
    * 命令行传参方式使用：执行 node `index.js --help` 查看。

#### 示例1 test.xlsx
| id   |              |        | []      | []          |   []       | {}           | [{}]                          |
| id   | desc         | flag   | nums    | words       |   map      | data         | hero                          |
| ---- | -------------| ------ | ------- | ----------- | ---------- | ------------ | --------------------------    |
| 123  | description  | true   | 1,2     | 哈哈,呵呵     | true,true  | a:123;b:45   | id:2;level:30,id:3;level:80  |
| 456  | 描述          | false  | 3,5,8   | shit,my god | false,true | a:11;b:22    | id:9;level:38,id:17;level:100 |


输出如下：

```json
[{
    "id": 123,
    "desc": "description",
    "flag": true,
    "nums": [1, 2],
    "words": ["哈哈", "呵呵"],
    "map": [true, true],
    "data": {
        "a": 123,
        "b": 45
    },
    "hero": [
      {"id": 2,"level": 30},
      {"id": 3,"level": 80}
    ]
}, {
    "id": 456,
    "desc": "描述",
    "flag": false,
    "nums": [3, 5, 8],
    "words": ["shit", "my god"],
    "map": [false, true],
    "data": {
        "a": 11,
        "b": 22
    },
    "hero": [
      {"id": 9, "level": 38 },
      {"id": 17,"level": 100}
    ]
}]
```

## 支持以下数据类型
* number 数字类型
* boolean  布尔
* string 字符串
* date 日期类型
* object 对象，复杂的嵌套可以通过外键来实现，见“外键类型的sheet关联”
* number-array  数字数组
* boolean-array  布尔数组
* string-array  字符串数组
* object-array 对象数组，复杂的嵌套可以通过外键来实现，见“外键类型的sheet关联”

## 表头规则
* 默认表格第一行为说明,第二行为数据类型,第三行为列名 。
* 默认数据类型为string
* 字符串类型：此列表头的命名形式 `string` 。
* 基本数据类型(string,number,bool)时候，一般不需要设置会自动判断，但是也可以明确声明数据类型。
* 字符串类型：此列表头的命名形式 `string` 。
* 数字类型：此列表头的命名形式 `number` 。
* 日期类型：`date` 。格式`YYYY/M/D H:m:s` or `YYYY/M/D` or `YYYY-M-D H:m:s` or `YYYY-M-D`。（==注意：目前xlsx文件里面列必须设置为文本类型，如果是日期类型的话，会导致底层插件解析出来错误格式的时间==）.
* 布尔类型：此列表头的命名形式 `bool` 。
* 基本类型数组：此列表头的命名形式 `[]` 。
* 对象：此列表头的命名形式 `{}` 。
* 对象数组：此列表头的命名形式`[{}]` 。
* id：此列表头的命名形式`id`，用来生成对象格式的输出，以该列字段作为key，一个sheet中不能存在多个id类型的列，否则会被覆盖，相关用例请查看test/heroes.xlsx
* id[]：此列表头的命名形式`id[]`，用来约束输出的值为对象数组，相关用例请查看test/stages.xlsx

## 数据规则
* 关键符号都是半角符号。
* 数组使用逗号`,`分割。
* 对象属性使用分号`;`分割。
* 列格式如果是日期，导出来的是格林尼治时间不是当时时区的时间，列设置成字符串可解决此问题。

## 外键类型的sheet关联
* sheet名称必须为【sheet.属性名称】，例如存在一个名称为a的sheet，会导出一个a.json，可以使用一个名称为a.b的sheet为这个json添加一个b的属性
* 外键类型的sheet（sub sheet）与被关联的sheet（master sheet）顺序上无限制
* master sheet的输出类型如果为对象，则sub sheet必须也存在master sheet同列名并且类型为id的列作为关联关系；master sheet的输出类型如果为数组，则sub sheet按照数组下标（行数）顺序关联
* 相关用例请查看test/heroes.xlsx

## 原理说明
* 依赖 `node-xlsx` 这个项目解析xlsx文件。
* xlsx就是个zip文件，解压出来都是xml。有一个xml存的string，有相应个xml存的sheet。通过解析xml解析出excel数据(json格式)，这个就是`node-xlsx` 做的工作。
* 本项目只需利用 `node-xlsx` 解析xlsx文件，然后拼装自己的json数据格式。

## 补充
* windows/mac/linux都支持。
* 项目地址 [xlsx2json master](https://github.com/koalaylj/xlsx2json)

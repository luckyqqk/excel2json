# excel2json
export excel data to json . the toolkit code by nodeJs v6.2.0, ES6 style.
### 感谢
* 感谢于小懒先生,初始项目down于他的[git](https://github.com/koalaylj/xlsx2json)
* 整体代码结构和思路均无变化,更改内容如下:
1. 将表头中的列名和数据类型分离成两行(根据我们的使用习惯,有时候列名过长,若类型掺在其中,影响表的阅读)
2. 重写了读写逻辑,去掉了配表中的页签顺序限制(原来属签一定要排在本签右边).
3. 增加了默认配置和默认数据(没填数据类型的字段(id类型必须填),默认为string类型,默认值为'').
4. 支持json数据直接转换(大多数情况不需要此功能)
5. 支持多页签导入同一json,如:页签drop和页签drop-1均会导入drop.json中.页签drop.out和页签drop-1.out均会导入drop.json中的out字段.

### 使用说明
* 只支持.xlsx格式,不支持.xls格式
* 本项目基于node编写,使用需先安装nodeJS6.0+
* 配置config.json
```json
{
    "xlsx": {
        "type": 2,                              // 表格第二行为数据类型行(基本就用id,number,string,[])
        "head": 3,                              // 第三行为表头所在的行，也是导出json的key。
        "excelDir": "",                         // xlsx文件夹,默认为本项目中的excel文件夹
        "outDir": "./json",                     // 导出的json存放的位置,支持多产出地址用','号分隔
        "arraySeparator":","                    // 数组的分隔符
    }
}
```
* 执行`export.sh/export.bat`即可将`excelDir` 文件导成json并存放到 `outDir` 下.json名字以excel的sheet名字命名.
#### 示例1 test.xlsx
| id   |              |        | []      | []          |   []       | {}           | [{}]                          |
|:----:|:------------:|:------:|:-------:|:-----------:|:----------:|:------------:|:-----------------------------:|
| id   | desc         | flag   | nums    | words       |   map      | data         | hero                          |
| 123  | description  | true   | 1,2     | 哈哈,呵呵     | true,true  | a:123;b:45   | id:2;level:30,id:3;level:80  |
| 456  | 描述          | false  | 3,5,8   | shit,my god | false,true | a:11;b:22    | id:9;level:38,id:17;level:100 |
* 输出如下：

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
### 支持以下数据类型
* number 数字类型
* boolean  布尔
* string 字符串
* date 日期类型
* object 对象,复杂的嵌套可以通过外键来实现,见“外键类型的sheet关联”
* number-array  数字数组
* boolean-array  布尔数组
* string-array  字符串数组
* object-array 对象数组,复杂的嵌套可以通过外键来实现,见“外键类型的sheet关联”

### 表头规则
* 默认表格第一行为说明,第二行为数据类型,第三行为列名
* 字符串类型：(默认) 此列表头的命名形式 `string`
* 数字类型：此列表头的命名形式 `number` 
* 基本类型数组：此列表头的命名形式 `[]`
* 对象：此列表头的命名形式 `{}`
* 对象数组：此列表头的命名形式`[{}]`
* id：此列表头的命名形式`id`,用来生成对象格式的输出,以该列字段作为key,一个sheet中不能存在多个id类型的列,否则会被覆盖,相关用例请查看test/heroes.xlsx
* id[]：此列表头的命名形式`id[]`,用来约束输出的值为对象数组,相关用例请查看test/stages.xlsx

### 数据规则
* 关键符号都是半角符号.
* 数组使用逗号`,`分割,支持自定义.
* 对象属性使用分号`;`分割.
* 列格式如果是日期,导出来的是格林尼治时间不是当时时区的时间,列设置成字符串可解决此问题.

### sheet(页签)导出
* sheet名称为即将导出的json文件名,例如存在一个名称为a的sheet,会导出一个a.json,且称之为:本签
* 若同sheet内容过长,或者在内容上区分可以分别管理,可以另立sheet,用'-'+数字命名即可.称之为:附签.如drop和drop-1,均会导出到drop.json

### 外键类型的sheet关联
* 外键类型的sheet命名必须为[sheet.属性名],称之为子签.如drop.out即可为drop的数据增加一个out的属性.支持drop-1.out 
* 子签与本签在顺序上无限制,但出于制表习惯,子签一般在本签右边
* 本签支持导出对象和数组.
* 若本签导出的是数组,则该本签不具有子签.
* 若本签导出的是对象,需有id列,类型也为id;子签也具有id列,类型为[].这样本签id与子签id即形成了一对多的关联关系
* 相关用例请查看test/heroes.xlsx

### 原理说明
* 依赖 `node-xlsx` 这个项目解析xlsx文件.
* xlsx就是个zip文件,解压出来都是xml.有一个xml存的string,有相应个xml存的sheet.通过解析xml解析出excel数据(json格式),这个就是`node-xlsx` 做的工作.
* 本项目只需利用 `node-xlsx` 解析xlsx文件,然后拼装自己的json数据格式.

### 系统支持
* windows/mac/linux

# OfficeOpenXml.Extension.AspNetCore

OfficeOpenXml.Extension.AspNetCore 是一个基于 OfficeOpenXml 拓展，它依赖于 [EPPlus](https://www.nuget.org/packages/EPPlus/5.0.3)，用于根据模板输出 Excel。

**注意：** 由于 Excel 2003 版本 和 2007 之后版本文件结构的差异性，当前扩展无法同时兼容两种模式，仅支持 *.xlsx 文件！！！

## 快速使用

### 1. 安装组件 

* [OfficeOpenXml.Extension.AspNetCore](https://www.nuget.org/packages/OfficeOpenXml.Extension.AspNetCore)

``` bash
dotnet add package OfficeOpenXml.Extension.AspNetCore
```

### 2.使用组件

#### 2.1 读取 Excel 模板， 导入数据

* 准备Excel模板
  * [Projects.xlsx](./samples/OfficeOpenXmlSample/wwwroot/templates/Projects.xlsx)

* 定义接收对象

  ```csharp
  [Worksheet(Index = 1, HasHeader = false)]
  public class ProjectRow
  {
      [Column(Number = 1)]
      public int Id { get; set; }
  
      [Column(Number = 2)]
      public string Name { get; set; }
  
      [Column(Number = 3)]
      public string Description { get; set; }
  }
  ```

  + Worksheet - 表格属性，其中 `Index` 对应表格的索引（从 0 开始），`HasHeader` 对应当前表格是否包含表头
  + Column - 单元格属性，其中 `Number` 对应单元格的列

* 读取 Excel 信息

  ```csharp
  [HttpGet("projects", Name = "Projects")]
  public IEnumerable<ProjectRow> GetProjects()
  {
      var excelFilePath = Path.Combine(_wwwroot, "templates", "Projects.xlsx");
      var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
      using var excelPackage = new ExcelPackage(fileStream);
      return excelPackage.ParseWorksheet<ProjectRow>().ToList();
  }
  ```

* 最终结果展示

  ```json
  [
    {
      "id": 1,
      "name": "MyHRW",
      "description": "Case Management Tool"
    },
    {
      "id": 2,
      "name": "PEX",
      "description": "Global Payroll Exchange"
    }
  ]
  ```

  

#### 2.2 读取 Excel 模板，导出数据

* 准备Excel模板

  * [tpl.xlsx](./samples/OfficeOpenXmlSample/wwwroot/templates/tpl.xlsx)

* 读取模板文件

  ```csharp
  var excelFilePath = Path.Combine(_wwwroot, "templates", "tpl.xlsx");
  var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
  using var excelPackage = new ExcelPackage(fileStream);
  var workBook = excelPackage.Workbook;
  ```

* 构造填充对象

  ```csharp
  Dictionary<string, IEnumerable<string>> _marketLists = new()
  {
      { "水果", new string[] { "桃子", "李子", "香蕉", "梨" } },
      { "蔬菜", new string[] { "青菜", "土豆", "黄瓜", "啤酒" } }
  };
  
  //构造model
  var model = new
  {
      ProjectName = "灰太狼",
      Name = "Jeff",
      CreatedAt = DateTime.Now,
      BuyerName = "Bill",
      Cates = _marketLists.Select(m => new
      {
          Name = m.Key,
          Items = m.Value.Select(n => new
          {
              Name = n,
              Price = (decimal)random.Next(1, 100),
              Amount = random.Next(1, 100)
          })
      })
  };
  ```

* 填充数据对象

  ```csharp
  // 下面的FillModel就是 OfficeOpenXml.Extension.AspNetCore 提供的拓展方法
  workBook.Worksheets.First().FillModel(model);
  ```

* 导出模板文件

  ```csharp
  string fileName = "Lists_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
  string exportFilePath = Path.Combine(_wwwroot, "outputs", fileName);
  var exportFile = new FileInfo(exportFilePath);
  excelPackage.SaveAs(exportFile);
  return File(exportFile.OpenRead(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
  ```

#### 2.3 其他功能辅助说明

  * 输出内容目前仅支持基础的变量、成员，不支持方法、运算等高级特性；控制代码目前仅支持 for 循环、嵌套 for 循环以及索引，使用索引时需要注意索引计数从1开始，因为excel中通常序号从1开始。
  * 输出公式的功能用 @= 开头便于程序识别，解析时会将 @ 去掉，后面的内容对 {...} 进行解释并替换。R[-4]表示相对值-4行，R[-1]表示相对值-1行，C后面没有 [] 表示当前列。
  * 具体细节可以进一步参考案例代码：
    * [ConsoleApp1](./samples/net45/Program.cs)
    * [OfficeOpenXmlSample](./samples/OfficeOpenXmlSample/Controllers/ExcelController.cs)

### 鸣谢

* [OfficeOpenXml.Entends 根据模板导出Excel](https://www.cnblogs.com/mhsg/p/7125112.html)

* [OfficeOpenXml.Extensions EPPlus Extensions](https://github.com/olivierl/OfficeOpenXml.Extensions)
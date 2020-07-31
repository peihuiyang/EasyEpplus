# EasyEpplus
> * 对EPPlus的简易封装，实现Excel的导入导出。
> * 本类库适合部署到Linux或者Windows上的项目。

## v1-0-0 2020-07-31

### 引用Epplus

* 使用NuGet引入包`Epplus`

* 在项目配置文件引入许可，如.NET Core WebApi项目

  ```
  "EPPlus": {
      "ExcelPackage": {
        "LicenseContext": "Commercial" //The license context used
      }
    }
  ```

  

### 文件说明

* EPPlusHelper：数据转换类
* ExcelExport：Excel文件导出实现类
* ExcelImport：Excel文件导入将数据转化为List<T>
* ExcelExportAttribute：Excel导出工作簿特性
* ExportColumnAttribute：Excel导出工作簿数据列特性

### 功能实现说明

* 导出：

  - 根据数据List<T>按照实体Dto及设置的特性导出数据

  - 导出为byte[]数据流，可以再加工进行输出
* 导入：接受Excel文件导入转化的Stream文件流，转化为List<T>类型数据进行加工存储
  
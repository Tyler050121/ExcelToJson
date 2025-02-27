# Excel 转 JSON 工具

一个高效的 Excel 表格转 JSON 数据格式转换工具，支持复杂数据结构及多种数据类型。

## 项目简介

这个工具可以将 Excel 表格数据转换为结构化的 JSON 数据，支持数组和对象两种输出格式，适用于游戏开发中的配置表转换、数据迁移等场景。

## 功能特点

- **多种输出格式**：支持数组格式（适用于列表数据）和对象格式（适用于字典数据）
- **丰富的数据类型**：支持基本类型（数值、字符串、布尔值）和复杂类型（数组、字典、自定义类）
- **良好的错误处理**：详细的错误提示和彩色日志输出
- **批量处理能力**：一次处理多个 Excel 文件
- **灵活的配置**：支持自定义输入和输出目录

## 使用方法

### 直接运行

```bash
cd ExcelToJson/Source
dotnet build
dotnet run
```

### 指定目录运行

```bash
dotnet run <Excel文件目录> <JSON输出目录>
```

## Excel 文件格式要求

Excel 文件需要按以下格式组织：

1. **第1行**：注释说明（可选）
2. **第2行**：数据类型定义
   - 第1列必须为 `array` 或 `object`
3. **第3行**：变量名（属性名）
4. **第4行及以后**：实际数据
   - 第1列非空的行会被识别为新对象

### 数据类型支持

#### 基本类型
- `number` - 自动检测整数或浮点数
- `int/integer` - 整数
- `long` - 长整数
- `float` - 单精度浮点数
- `double` - 双精度浮点数
- `bool/boolean` - 布尔值（支持 1/0、t/f、true/false）
- `string` - 字符串

#### 复杂类型
- `array` - 数组类型
- `dict` - 字典类型
- `classN` - 类对象（N 表示属性数量）

## 数据格式示例

### 数组格式（array）

适用于生成数组列表数据：(第一行注释忽略)

| array | number | string |   bool  |
|-------|--------|--------|---------|
|       | id     | name   |  active |      
| 1     | 1      | 物品1   |  true   |      
| 2     | 2      | 物品2   |  false  |

生成的 JSON:
```json
[
  {
    "name": "物品1",
    "active": true
  },
  {
    "name": "物品2",
    "active": false
  }
]
```

### 对象格式（object）

适用于生成字典数据：(第一行注释忽略)

| object | int | string | number |
|--------|--------|--------|--------|
|        | key    | name   | value  |
| 1      | 1001   | 物品1  | 100    |
| 2      | 1002   | 物品2  | 200    |

生成的 JSON:
```json
{
  "1001": {
    "name": "物品1",
    "value": 100
  },
  "1002": {
    "name": "物品2",
    "value": 200
  }
}
```

## 日志系统

程序提供了五种级别的彩色日志输出：

- **Info**（白色）：普通信息
- **Warning**（黄色）：警告信息
- **Error**（红色）：错误信息
- **Success**（绿色）：成功信息
- **Tip**（青色）：提示信息

## 项目结构

```
ExcelToJson/
├── .gitignore           # Git 忽略文件配置
├── README.md            # 项目说明文档
├── Source/         # 源代码目录
│   ├── Program.cs       # 主程序入口
│   ├── ExcelProcessor.cs # Excel 处理核心类
│   ├── Logger.cs        # 日志处理类
│   └── ExcelToJson.csproj # 项目配置文件
├── Tables/              # Excel 文件目录
│   └── TestConfig.xlsx  # 示例 Excel 文件
└── Output/              # JSON 输出目录
```

## 运行环境

- .NET 9.0 或更高版本
- 依赖包：
  - EPPlus 7.6.0（Excel 文件处理）
  - Newtonsoft.Json 13.0.3（JSON 序列化）

## 注意事项

1. Excel 文件可以是 `.xlsx` 或 `.xlsm` 格式
2. 表格中的空单元格会被自动忽略
3. 转换后的 JSON 文件会保存在输出目录，文件名与 Excel 文件同名
4. 字典类型的键如果重复，后面的键值将覆盖前面的值
using System.Text.RegularExpressions;
using OfficeOpenXml;

public class ExcelProcessor
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<string> _dataTypes;
    private readonly List<string> _variableNames;
    private readonly List<int> _newObjectRows;

    enum JsonType
    {
        Array,
        Object
    }

    public ExcelProcessor(string excelFile)
    {
        // 设置EPPlus的LicenseContext
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var package = new ExcelPackage(new FileInfo(excelFile));
        _worksheet = package.Workbook.Worksheets[0];
        (_dataTypes, _variableNames, _newObjectRows) = ReadExcelHeader(_worksheet);
    }

    public object? GetResult()
    {
        var jsonType = GetJsonType(_worksheet);
        var totalColumns = _worksheet.Dimension.End.Column;

        if (jsonType == JsonType.Array)
        {
            return ProcessArrayType(totalColumns);
        }

        if (jsonType == JsonType.Object)
        {
            return ProcessObjectType(totalColumns);
        }

        return null;
    }

    private object ProcessArrayType(int totalColumns)
    {
        var data = new List<object>();

        foreach (var currentRow in _newObjectRows)
        {
            var rowData = new Dictionary<string, object>();
            var currentColumn = 2;

            while (currentColumn <= totalColumns)
            {
                var (value, columnsRead) = ConvertValue(currentRow, currentColumn);
                if (value != null)
                {
                    rowData[_variableNames[currentColumn - 1]] = value;
                }
                currentColumn += columnsRead;
            }

            if (rowData.Count > 0)
            {
                data.Add(rowData);
            }
        }

        return data;
    }

    private object ProcessObjectType(int totalColumns)
    {
        var result = new Dictionary<object, object>();

        foreach (var currentRow in _newObjectRows)
        {
            var (key, _) = ConvertValue(currentRow, 2);
            if (key != null)
            {
                var rowData = new Dictionary<string, object>();
                var currentColumn = 3;
                while (currentColumn <= totalColumns)
                {
                    var propName = _variableNames[currentColumn - 1];
                    if (string.IsNullOrEmpty(propName))
                    {
                        throw new FormatException($"属性名为空 [列:{currentColumn}]");
                    }
                    
                    var (value, columnsRead) = ConvertValue(currentRow, currentColumn);
                    if (value != null)
                    {
                        rowData[propName] = value;
                    }
                    currentColumn += columnsRead;
                }

                if (rowData.Count > 0)
                {
                    if (result.ContainsKey(key))
                    {
                        Logger.Log(LogLevel.Warning, $"第 {currentRow} 行的键 '{key}' 重复，将被覆盖");
                    }
                    result[key] = rowData;
                }
            }
        }

        return result;
    }

    private (List<string> dataTypes, List<string> variableNames, List<int> newObjectRows)
        ReadExcelHeader(ExcelWorksheet worksheet)
    {
        var dataTypes = new List<string>();
        var variableNames = new List<string>();
        var newObjectRows = new List<int>();

        var totalColumns = worksheet.Dimension.End.Column;
        var totalRows = worksheet.Dimension.End.Row;
        
        Logger.Log(LogLevel.Tip, $"总列数: {totalColumns}, 总行数: {totalRows}, 若报错请将空列删除");

        if (totalRows < 4)  // 至少需要4行：注释、类型、变量名、数据
        {
            throw new FormatException("Excel格式错误: 数据行数不足");
        }

        // 读取数据类型和变量名
        for (var currentColumn = 1; currentColumn <= totalColumns; currentColumn++)
        {
            var dataType = worksheet.Cells[2, currentColumn].Text.ToLower();
            var variableName = worksheet.Cells[3, currentColumn].Text;

            if (string.IsNullOrEmpty(dataType))
            {
                throw new FormatException($"第2行第{currentColumn}列的数据类型为空");
            }

            dataTypes.Add(dataType);
            variableNames.Add(variableName);
        }

        // 查找数据行
        for (var currentRow = 4; currentRow <= totalRows; currentRow++)
        {
            if (!string.IsNullOrWhiteSpace(worksheet.Cells[currentRow, 1].Text))
            {
                newObjectRows.Add(currentRow);
            }
        }

        if (newObjectRows.Count == 0)
        {
            throw new FormatException("未找到有效的数据行");
        }

        return (dataTypes, variableNames, newObjectRows);
    }

    private JsonType GetJsonType(ExcelWorksheet worksheet)
    {
        string type = worksheet.Cells[2, 1].Text.ToLower();
        return type switch
        {
            "array" => JsonType.Array,
            "object" => JsonType.Object,
            _ => throw new FormatException("第一列第二行必须为array或object")
        };
    }

    private List<int> GetNewObjectRows(int row, int column)
    {
        var totalRows = _worksheet.Dimension.End.Row;
        var newObjectRows = new List<int>();
        var foundFirstOne = false;

        if (string.IsNullOrWhiteSpace(_worksheet.Cells[row, column].Text))
        {
            // 第一格没有数据，直接返回
            return newObjectRows;
        }
        
        for (var currentRow = row; currentRow <= totalRows; currentRow++)
        {
            var cellValue = _worksheet.Cells[currentRow, column].Text;
            if (string.IsNullOrWhiteSpace(cellValue)) continue;
            if (cellValue == "1")
            {
                if (foundFirstOne) break;
                foundFirstOne = true;
            }

            newObjectRows.Add(currentRow);
        }

        return newObjectRows;
    }

    private object? ConvertDictionary(int row, int column)
    {
        var newObjectRows = GetNewObjectRows(row, column);
        var dictResult = new Dictionary<object, object>();
        foreach (var newRow in newObjectRows)
        {
            var (key, _) = ConvertValue(newRow, column + 1);
            if (key != null)
            {
                var currentColumn = column + 2;
                var (value, columnsRead) = ConvertValue(newRow, currentColumn);

                if (value != null)
                {
                    if (dictResult.ContainsKey(key))
                    {
                        Logger.Log(LogLevel.Warning, $"字典中的键 '{key}' 重复，将被覆盖");
                    }
                    dictResult[key] = value;
                }
            }
        }

        return dictResult.Count == 0 ? null : dictResult;
    }

    private object? ConvertArray(int row, int column)
    {
        var newObjectRows = GetNewObjectRows(row, column);
        var arrayResult = new List<object>();
        foreach (var newRow in newObjectRows)
        {
            var (value, columnsRead) = ConvertValue(newRow, column + 1);
            if (value != null)
            {
                arrayResult.Add(value);
            }
        }

        return arrayResult.Count == 0 ? null : arrayResult.ToArray();
    }

    private object? ConvertClass(int row, int column)
    {
        var match = Regex.Match(_dataTypes[column - 1].ToLower(), @"class(\d+)");
        if (!match.Success)
        {
            throw new FormatException($"无效的类型格式: {_dataTypes[column - 1]}，应为class+数字 (如class2)");
        }

        var propertyCount = int.Parse(match.Groups[1].Value);
        var classResult = new Dictionary<string, object>();
        var currentColumn = column + 1;

        for (var i = 0; i < propertyCount; i++)
        {
            var propName = _variableNames[currentColumn - 1];
            if (string.IsNullOrEmpty(propName))
            {
                throw new FormatException($"类的属性名为空 [列:{currentColumn}]");
            }

            var (value, columnsRead) = ConvertValue(row, currentColumn);
            if (value != null)
            {
                classResult[propName] = value;
            }

            currentColumn += columnsRead;
        }

        return classResult.Count == 0 ? null : classResult;
    }

    private (object? value, int columnsRead) ConvertValue(int row, int column)
    {
        var dataType = _dataTypes[column - 1];

        if (string.IsNullOrEmpty(dataType))
        {
            throw new FormatException($"数据类型为空 [行:{row}, 列:{column}]");
        }

        if (dataType.StartsWith("arr"))
        {
            return (ConvertArray(row, column), CalculateSkippedColumns(column));
        }

        if (dataType.StartsWith("dict"))
        {
            return (ConvertDictionary(row, column), CalculateSkippedColumns(column));
        }

        if (dataType.StartsWith("class"))
        {
            return (ConvertClass(row, column), CalculateSkippedColumns(column));
        }

        return (ConvertBasicValue(dataType, row, column), 1);
    }

    protected virtual object? ConvertBasicValue(string dataType, int row, int column)
    {
        var data = _worksheet.Cells[row, column].Text;

        if (string.IsNullOrWhiteSpace(data))
        {
            Logger.Log(LogLevel.Warning, $"单元格值为空 [行:{row}, 列:{column}]");
            return null;
        }

        return dataType.ToLower() switch
        {
            "number" => data.Contains(".") ? ParseFloat(data, row, column) : ParseInt(data, row, column),
            "int" or "integer" => ParseInt(data, row, column),
            "long" => ParseLong(data, row, column),
            "float" => ParseFloat(data, row, column),
            "double" => ParseDouble(data, row, column),
            "bool" or "boolean" => ParseBool(data, row, column),
            "string" => data,
            _ => throw new FormatException($"不支持的数据类型: {dataType}")
        };
    }

    private int ParseInt(string value, int row, int column)
    {
        if (!int.TryParse(value, out int result))
        {
            throw new FormatException($"无法将值 '{value}' 转换为整数 [行:{row}, 列:{column}]");
        }
        return result;
    }

    private long ParseLong(string value, int row, int column)
    {
        if (!long.TryParse(value, out long result))
        {
            throw new FormatException($"无法将值 '{value}' 转换为长整数 [行:{row}, 列:{column}]");
        }
        return result;
    }

    private float ParseFloat(string value, int row, int column)
    {
        if (!float.TryParse(value, out float result))
        {
            throw new FormatException($"无法将值 '{value}' 转换为浮点数 [行:{row}, 列:{column}]");
        }
        return result;
    }

    private double ParseDouble(string value, int row, int column)
    {
        if (!double.TryParse(value, out double result))
        {
            throw new FormatException($"无法将值 '{value}' 转换为双精度浮点数 [行:{row}, 列:{column}]");
        }
        return result;
    }

    private bool ParseBool(string value, int row, int column)
    {
        value = value.Trim().ToLower();
        if (value == "1" || value == "t" || value == "true") return true;
        if (value == "0" || value == "f" || value == "false") return false;

        throw new FormatException($"无效的布尔值: '{value}' [行:{row}, 列:{column}]");
    }

    private int CalculateSkippedColumns(int column)
    {
        var dataType = _dataTypes[column - 1].ToLower();

        if (string.IsNullOrEmpty(dataType))
        {
            throw new FormatException($"数据类型为空 [列:{column}]");
        }

        if (dataType.StartsWith("dict"))
        {
            var valueColumn = column + 2;
            return 2 + CalculateSkippedColumns(valueColumn);
        }

        if (dataType.StartsWith("array"))
        {
            var valueColumn = column + 1;
            return 1 + CalculateSkippedColumns(valueColumn);
        }

        if (dataType.StartsWith("class"))
        {
            var match = Regex.Match(dataType, @"class(\d+)");
            if (!match.Success)
            {
                throw new FormatException($"无效的类型格式: {dataType}，应为class+数字 (如class2)");
            }

            var propertyCount = int.Parse(match.Groups[1].Value);
            var propSkip = 0;
            for (var i = 1; i <= propertyCount; i++)
            {
                propSkip += CalculateSkippedColumns(column + i);
            }

            return propSkip + 1;
        }

        return 1;
    }
}
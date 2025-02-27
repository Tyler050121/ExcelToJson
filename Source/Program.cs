using Newtonsoft.Json;

public static class Program
{
    private static void Main(string[] args)
    {
        try
        {
            // 获取输出目录路径（从命令行参数或使用默认值）
            var outputDir = args.Length > 0 ? args[1] : "../Output";
            Directory.CreateDirectory(outputDir);

            // 获取Excel文件目录路径（从命令行参数或使用默认值）
            var excelDir = args.Length > 0 ? args[0] : "../Tables";
            if (!Directory.Exists(excelDir))
            {
                Logger.Log(LogLevel.Error, $"目录不存在: {excelDir}");
                return;
            }

            // 获取指定目录下所有Excel文件
            // 获取指定目录下所有Excel文件
            var excelFiles = Directory.GetFiles(excelDir, "*.*")
                .Where(file => (file.EndsWith(".xlsm") || file.EndsWith(".xlsx")) && 
                               !Path.GetFileName(file).StartsWith(".~") && 
                               !Path.GetFileName(file).StartsWith("~$"))
                .ToArray();

            if (excelFiles.Length == 0)
            {
                Logger.Log(LogLevel.Warning, "没有Excel文件需要处理");
                return;
            }

            Logger.Log(LogLevel.Tip, $"在 {excelDir} 中找到 {excelFiles.Length} 个Excel文件需要处理\n");

            var hasError = false;

            // 处理excel文件
            foreach (var excelFile in excelFiles)
            {
                try
                {
                    ProcessExcelFile(excelFile, outputDir);
                }
                catch (Exception ex)
                {
                    Logger.Log(LogLevel.Error, $"处理失败: {ex.Message}\n");
                    hasError = true;
                }
            }

            Logger.Log(LogLevel.Tip, $"输出目录: {Path.GetFullPath(outputDir)}");
            if (hasError)
            {
                Logger.Log(LogLevel.Error, "部分文件处理失败, 请查看日志");
            }
            else
            {
                Logger.Log(LogLevel.Success, "所有文件处理完成!");
            }
        }
        catch (Exception ex)
        {
            Logger.Log(LogLevel.Error, $"程序执行出错: {ex.Message}");
        }
    }

    private static void ProcessExcelFile(string excelFile, string outputDir)
    {
        Logger.Log(LogLevel.Info, $"正在处理: {Path.GetFileName(excelFile)}");

        var processor = new ExcelProcessor(excelFile);
        var result = processor.GetResult();

        var jsonFileName = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(excelFile) + ".json");
        if (result != null)
        {
            SaveJsonToFile(result, jsonFileName);
            Logger.Log(LogLevel.Info, $"已保存到: {jsonFileName}");
            Logger.Log(LogLevel.Info, "处理成功\n");
        }
        else
        {
            Logger.Log(LogLevel.Error, "处理失败\n");
        }
    }

    private static void SaveJsonToFile(object result, string jsonFileName)
    {
        var settings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore
        };

        var json = JsonConvert.SerializeObject(result, settings);
        File.WriteAllText(jsonFileName, json);
    }
}
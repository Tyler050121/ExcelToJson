internal enum LogLevel
{
    Info,
    Warning,
    Error,
    Success,
    Tip
}

internal static class Logger
{

    public static void Log(LogLevel level, string message)
    {
        var color = level switch
        {
            LogLevel.Info => ConsoleColor.White,
            LogLevel.Warning => ConsoleColor.Yellow,
            LogLevel.Error => ConsoleColor.Red,
            LogLevel.Success => ConsoleColor.Green,
            LogLevel.Tip => ConsoleColor.Cyan,
            _ => ConsoleColor.White
        };

        var originalColor = Console.ForegroundColor;
        Console.ForegroundColor = color;
        Console.WriteLine($"[{level}]: {message}");
        
        Console.ForegroundColor = originalColor;
    }
}
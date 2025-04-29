namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class Logger
{
    public static void Info(string message)
    {
        Write(message, ConsoleColor.Gray);
    }

    public static void Success(string message)
    {
        Write(message, ConsoleColor.Green);
    }

    public static void Warn(string message)
    {
        Write("⚠️ " + message, ConsoleColor.Yellow);
    }

    public static void Error(string message)
    {
        Write("❌ " + message, ConsoleColor.Red);
    }

    private static void Write(string message, ConsoleColor color)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        Console.ForegroundColor = color;
        Console.WriteLine($"[{timestamp}] {message}");
        Console.ResetColor();
    }
}

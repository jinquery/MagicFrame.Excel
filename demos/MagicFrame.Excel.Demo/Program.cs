using MagicFrame.Excel.Demo;

namespace MagicFrame.Excel.Demo;

/// <summary>
/// MagicFrame.Excel Demo 入口
/// 用法：
///   dotnet run --project demos/MagicFrame.Excel.Demo            # 全部演示
///   dotnet run --project demos/MagicFrame.Excel.Demo -- basic   # 仅基础演示
///   dotnet run --project demos/MagicFrame.Excel.Demo -- ext     # 仅扩展演示
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "";
        string outputDir = args.Length > 1
            ? args[1]
            : Path.Combine(AppContext.BaseDirectory, "output");

        Console.WriteLine("MagicFrame.Excel Demo");
        Console.WriteLine("======================");
        Console.WriteLine();

        switch (mode)
        {
            case "basic":
                BasicDemo.Run(outputDir);
                break;
            case "ext":
            case "extension":
                ExtensionDemo.Run(outputDir);
                break;
            default:
                BasicDemo.Run(outputDir);
                ExtensionDemo.Run(outputDir);
                break;
        }

        Console.WriteLine($"输出目录: {outputDir}");
        if (!Console.IsInputRedirected)
        {
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }
}

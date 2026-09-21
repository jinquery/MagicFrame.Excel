namespace MagicFrame.Excel.Exceptions;

/// <summary>
/// Excel 导入/处理异常
/// </summary>
public class ExcelImportException : Exception
{
    public ExcelImportException() : base("Excel 处理异常") { }

    public ExcelImportException(string message) : base(message) { }

    public ExcelImportException(string message, Exception inner) : base(message, inner) { }
}

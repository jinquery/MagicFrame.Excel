namespace MagicFrame.Excel.Options;

/// <summary>
/// 导入过程中的结构化问题（E4）：比 <see cref="ExcelImportOptions.Errors"/> 的纯文本更便于程序化处理。
/// </summary>
public sealed class ImportIssue
{
    /// <summary>严重级别</summary>
    public ImportIssueSeverity Severity { get; }

    /// <summary>机器可读的错误码，如 MissingColumn / UnexpectedColumn / RowWriteFailed</summary>
    public string Code { get; }

    /// <summary>人类可读描述</summary>
    public string Message { get; }

    /// <summary>相关行号（1 基，无则 null）</summary>
    public int? RowNumber { get; }

    private ImportIssue(ImportIssueSeverity severity, string code, string message, int? rowNumber)
    {
        Severity = severity;
        Code = code;
        Message = message;
        RowNumber = rowNumber;
    }

    public static ImportIssue Info(string code, string message, int? row = null)
        => new(ImportIssueSeverity.Info, code, message, row);

    public static ImportIssue Warning(string code, string message, int? row = null)
        => new(ImportIssueSeverity.Warning, code, message, row);

    public static ImportIssue Error(string code, string message, int? row = null)
        => new(ImportIssueSeverity.Error, code, message, row);

    public override string ToString()
        => $"[{Severity}] {Code}{(RowNumber.HasValue ? $" @row{RowNumber}" : "")}: {Message}";
}

/// <summary>导入问题严重级别</summary>
public enum ImportIssueSeverity
{
    Info,
    Warning,
    Error
}

using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Engine;

/// <summary>
/// 导出结果：同时提供 工作簿/字节/Base64/流 多种形态。
/// <para>
/// Base64 惰性生成（仅访问时计算，避免只取 Bytes 的调用方多一次 1.3× 大小的字符串分配）。
/// 建议用完 <see cref="Dispose"/>（便捷方法 ExportBytes/ExportBase64/ExportToFile 已内部释放）。
/// </para>
/// </summary>
public class ExcelExportResult : IDisposable
{
    /// <summary>NPOI 工作簿</summary>
    public IWorkbook Workbook { get; }

    /// <summary>文件字节数组</summary>
    public byte[] Bytes { get; }

    /// <summary>Base64 字符串（供前端下载，惰性计算）</summary>
    public string Base64 => _base64 ??= Convert.ToBase64String(Bytes);

    /// <summary>内存流（Position 已归零）</summary>
    public MemoryStream Stream { get; }

    private string? _base64;

    internal ExcelExportResult(IWorkbook workbook, MemoryStream stream, byte[] bytes)
    {
        Workbook = workbook;
        Stream = stream;
        Bytes = bytes;
    }

    public void Dispose()
    {
        Stream?.Dispose();
        (Workbook as IDisposable)?.Dispose();
        GC.SuppressFinalize(this);
    }
}

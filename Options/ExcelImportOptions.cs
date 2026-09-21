using MagicFrame.Excel.Model;

namespace MagicFrame.Excel.Options;

/// <summary>
/// 导入选项
/// </summary>
public class ExcelImportOptions
{
    /// <summary>Sheet 名；为空时按 SheetIndex 取</summary>
    public string SheetName { get; set; } = "";

    /// <summary>Sheet 索引（SheetName 为空时生效），默认 0</summary>
    public int SheetIndex { get; set; }

    /// <summary>表头类型，默认单表头；必须与导出时一致</summary>
    public string HeaderKind { get; set; } = HeaderKinds.Single;

    /// <summary>数据起始行（1 基，指表头所在行，默认 1），数据从下一行开始</summary>
    public int DataStartRowIndex { get; set; } = 1;

    /// <summary>是否跳过全空行，默认 true</summary>
    public bool SkipEmptyRows { get; set; } = true;

    /// <summary>缺少必需列时是否抛异常，默认 true</summary>
    public bool ThrowOnMissingColumns { get; set; } = true;

    /// <summary>
    /// 发现 Excel 中多余的列（无法对应任何实体列）时是否抛异常，
    /// 默认 false（仅写入 <see cref="Errors"/> 提示）
    /// </summary>
    public bool ThrowOnUnexpectedColumns { get; set; }

    /// <summary>
    /// 是否忽略映射到只读/不存在属性的列：
    /// true 时不写、不视为缺失（默认）；false 时按必需列校验，缺失则抛异常。
    /// </summary>
    public bool IgnoreReadOnlyColumns { get; set; } = true;

    /// <summary>
    /// 逐行解析出错时是否收集到 <see cref="Issues"/> 而非直接抛出（默认 false=抛异常）。
    /// 开启后某行转换失败会被跳过并记录，其余行正常导入。
    /// </summary>
    public bool CollectRowErrors { get; set; }

    /// <summary>若设置，把每行对应的 Excel 行号（1 基）写回该实体属性</summary>
    public string? SourceRowProperty { get; set; }

    /// <summary>导入进度回调（0~1）</summary>
    public IProgress<double>? Progress { get; set; }

    /// <summary>列定义；为空时由列提供器从实体类型解析</summary>
    public IReadOnlyList<ExcelColumn>? Columns { get; set; }

    /// <summary>导入过程中的提示/错误信息（异常时也会写入）</summary>
    public IList<string> Errors { get; } = new List<string>();

    /// <summary>导入过程中的结构化问题（E4）</summary>
    public IList<ImportIssue> Issues { get; } = new List<ImportIssue>();

    /// <summary>
    /// 复制选项并指定 Sheet 名（供 <see cref="Engine.ExcelEngine.ImportAll{T}"/> 逐 Sheet 使用；
    /// Errors/Issues 为新集合，Columns 引用共享）
    /// </summary>
    internal ExcelImportOptions CopyForSheet(string sheetName)
    {
        return new ExcelImportOptions
        {
            SheetName = sheetName,
            SheetIndex = SheetIndex,
            HeaderKind = HeaderKind,
            DataStartRowIndex = DataStartRowIndex,
            SkipEmptyRows = SkipEmptyRows,
            ThrowOnMissingColumns = ThrowOnMissingColumns,
            ThrowOnUnexpectedColumns = ThrowOnUnexpectedColumns,
            IgnoreReadOnlyColumns = IgnoreReadOnlyColumns,
            CollectRowErrors = CollectRowErrors,
            SourceRowProperty = SourceRowProperty,
            Progress = Progress,
            Columns = Columns,
        };
    }
}

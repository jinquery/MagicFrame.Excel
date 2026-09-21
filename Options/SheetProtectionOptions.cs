namespace MagicFrame.Excel.Options;

/// <summary>
/// 工作表保护选项。
/// 说明：NPOI 的 ProtectSheet 默认会禁止 插入行/删除行/格式调整/排序/筛选 等操作，
/// 本类通过 CT_SheetProtection 显式放开，实现"锁定指定列的同时允许用户新增行"。
/// </summary>
public class SheetProtectionOptions
{
    /// <summary>是否启用保护；false 时即使存在锁定列也不保护</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>保护密码，为空时不保护；配合 <see cref="ProtectWithoutPassword"/> 可"保护但不设密码"</summary>
    public string? Password { get; set; } = "1234";

    /// <summary>置 true 且 Password 为空时：保护工作表但不设置密码</summary>
    public bool ProtectWithoutPassword { get; set; }

    /// <summary>允许用户新增行</summary>
    public bool AllowInsertRows { get; set; } = true;

    /// <summary>允许用户删除行</summary>
    public bool AllowDeleteRows { get; set; } = true;

    /// <summary>允许用户调整单元格/列/行格式（列宽、字号等）</summary>
    public bool AllowFormat { get; set; } = true;

    /// <summary>允许用户对受保护工作表进行排序</summary>
    public bool AllowSort { get; set; } = true;

    /// <summary>允许用户在受保护工作表上使用自动筛选</summary>
    public bool AllowAutoFilter { get; set; } = true;
}

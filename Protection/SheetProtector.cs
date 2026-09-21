using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MagicFrame.Excel.Protection;

/// <summary>
/// 工作表保护器：保护工作表并放开指定操作。
/// 说明：
/// - NPOI 的 ProtectSheet 默认会禁止 插入行/删除行/格式调整 等操作，
///   必须通过 CT_SheetProtection 显式放开，才能实现"锁定指定列的同时允许用户新增行"。
/// - CT_SheetProtection 的 sort / autoFilter 默认值为 true（表示禁止）：
///   若不显式写为 false，受保护工作表在 Excel 中 排序/自动筛选 会被置灰禁用。
/// - 空密码保护：设置 <see cref="SheetProtectionOptions.ProtectWithoutPassword"/> = true 且
///   Password 为空时，直接写 sheetProtection（不写 password 哈希），实现"保护但不设密码"。
/// </summary>
public static class SheetProtector
{
    /// <summary>
    /// 应用保护
    /// </summary>
    public static void Apply(ISheet sheet, SheetProtectionOptions options)
    {
        if (options == null || !options.Enabled)
        {
            return;
        }
        // Password 为 null 表示不保护；空密码仅当 ProtectWithoutPassword 时表示"保护但不设密码"
        if (options.Password == null || (options.Password.Length == 0 && !options.ProtectWithoutPassword))
        {
            return;
        }

        if (sheet is not XSSFSheet xssfSheet)
        {
            return;
        }

        if (!string.IsNullOrEmpty(options.Password))
        {
            sheet.ProtectSheet(options.Password);
        }
        else
        {
            // 无密码保护：ProtectSheet("") 会写入空密码哈希，不符合"无密码"语义，直接写 CT
            var ct = xssfSheet.GetCTWorksheet();
            var sp = ct.sheetProtection ?? ct.AddNewSheetProtection();
            sp.sheet = true;
            sp.objects = true;
            sp.scenarios = true;
        }

        var protection = xssfSheet.GetCTWorksheet()?.sheetProtection;
        if (protection == null)
        {
            return;
        }

        if (options.AllowInsertRows)
        {
            protection.insertRows = false;
        }
        if (options.AllowDeleteRows)
        {
            protection.deleteRows = false;
        }
        if (options.AllowFormat)
        {
            protection.formatCells = false;
            protection.formatColumns = false;
            protection.formatRows = false;
        }
        if (options.AllowSort)
        {
            protection.sort = false;
        }
        if (options.AllowAutoFilter)
        {
            protection.autoFilter = false;
        }
    }
}

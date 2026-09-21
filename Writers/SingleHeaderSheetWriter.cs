using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Writers;

/// <summary>
/// 单表头写入策略：第 1 行为表头（锁定），第 2 行起为数据
/// </summary>
public class SingleHeaderSheetWriter : BaseSheetWriter
{
    public override int HeaderRowCount => 1;

    protected override void WriteHeader(ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns, ExcelExportOptions options, WriterStyles styles)
    {
        IRow headerRow = sheet.CreateRow(0);
        for (int c = 0; c < columns.Count; c++)
        {
            ICell headerCell = headerRow.CreateCell(c);
            headerCell.SetCellValue(columns[c].Name);
            headerCell.CellStyle = styles.HeaderStyleFor(workbook, columns[c].HeaderColor);

            ApplyColumn(sheet, c, columns[c], options);
        }
    }
}

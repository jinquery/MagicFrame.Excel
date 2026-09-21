using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace MagicFrame.Excel.Writers;

/// <summary>
/// 分组表头写入策略（二级表头）：
/// 第 1 行为分组行，同一 GroupName 的多列合并显示 GroupText；
/// 第 2 行为各列列名；无分组的列上下合并两行显示列名。数据从第 3 行开始。
/// </summary>
public class GroupHeaderSheetWriter : BaseSheetWriter
{
    public override int HeaderRowCount => 2;

    protected override void WriteHeader(ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns, ExcelExportOptions options, WriterStyles styles)
    {
        IRow groupRow = sheet.CreateRow(0);
        IRow nameRow = sheet.CreateRow(1);

        int n = columns.Count;
        int c = 0;
        while (c < n)
        {
            var col = columns[c];

            if (string.IsNullOrWhiteSpace(col.GroupName))
            {
                // 无分组：上下合并两行显示列名
                sheet.AddMergedRegion(new CellRangeAddress(0, 1, c, c));
                ICell cell = groupRow.GetCell(c) ?? groupRow.CreateCell(c);
                cell.SetCellValue(col.Name);
                cell.CellStyle = styles.HeaderStyleFor(workbook, col.HeaderColor);
                ApplyColumn(sheet, c, col, options);
                c++;
                continue;
            }

            // 分组：第 1 行合并显示分组文字，第 2 行逐个显示列名
            string groupName = col.GroupName;
            int start = c;
            int end = c;
            while (end + 1 < n && columns[end + 1].GroupName == groupName)
            {
                end++;
            }

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, start, end));
            ICell groupCell = groupRow.GetCell(start) ?? groupRow.CreateCell(start);
            groupCell.SetCellValue(columns[start].GroupText);
            groupCell.CellStyle = styles.HeaderStyleFor(workbook, columns[start].HeaderColor);

            for (int j = start; j <= end; j++)
            {
                ICell nameCell = nameRow.GetCell(j) ?? nameRow.CreateCell(j);
                nameCell.SetCellValue(columns[j].Name);
                nameCell.CellStyle = styles.HeaderStyleFor(workbook, columns[j].HeaderColor);
                ApplyColumn(sheet, j, columns[j], options);
            }

            c = end + 1;
        }
    }
}

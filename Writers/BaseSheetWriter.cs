using System.Text;
using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Converters;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using MagicFrame.Excel.Protection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace MagicFrame.Excel.Writers;

/// <summary>
/// Sheet 写入基类：承载列宽/隐藏列/整列锁定/公式/下拉/冻结/保护/筛选等公共逻辑。
/// 子类只需实现 <see cref="HeaderRowCount"/> 与 <see cref="WriteHeader"/>，即可得到不同的表头布局。
/// 本类无跨调用状态（样式按 Workbook 在每次写入时创建），可安全单例复用。
/// </summary>
public abstract class BaseSheetWriter : ISheetWriter
{
    /// <summary>表头占用行数</summary>
    public abstract int HeaderRowCount { get; }

    public void Write(
        ISheet sheet,
        IWorkbook workbook,
        IReadOnlyList<ExcelColumn> columns,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        IFormulaResolver formulaResolver,
        IEnumerable<object> rows,
        ExcelExportOptions options)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (columns == null) throw new ArgumentNullException(nameof(columns));

        // 样式属于当前工作簿，每次写入重新创建（避免跨工作簿复用样式）；边框作用于整表
        var styles = WriterStyles.Create(workbook, options.EnableBorders);

        // 整列锁定铺底需在写入单元格之前执行（NPOI 的 SetDefaultColumnStyle 会覆盖已存在单元格样式）
        ApplyColumnDefaults(workbook, sheet, columns);

        WriteHeader(sheet, workbook, columns, options, styles);

        WriteData(sheet, workbook, columns, accessor, converter, formulaResolver, rows, options, styles);

        ApplyValidation(sheet, columns, options);

        // 冻结窗格：冻结列 + 冻结表头可同时生效
        if (options.FreezeHeader || options.FreezeColumns > 0)
        {
            int cols = Math.Max(0, options.FreezeColumns);
            int rowsFrozen = options.FreezeHeader ? HeaderRowCount : 0;
            sheet.CreateFreezePane(cols, rowsFrozen);
        }

        // 自动筛选（作用于表头行）
        if (options.AutoFilter && columns.Count > 0)
        {
            sheet.SetAutoFilter(new CellRangeAddress(0, HeaderRowCount - 1, 0, columns.Count - 1));
        }

        ApplyProtection(sheet, columns, options);
    }

    /// <summary>子类实现：写入表头</summary>
    protected abstract void WriteHeader(ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns, ExcelExportOptions options, WriterStyles styles);

    // ---------- 整列锁定 ----------

    /// <summary>
    /// 按列的 IsLocked 对整列设置默认锁定/解锁样式，
    /// 保证用户新增行/新增单元格时，锁定列保持锁定、未锁定列可编辑。
    /// </summary>
    private static void ApplyColumnDefaults(IWorkbook workbook, ISheet sheet, IReadOnlyList<ExcelColumn> columns)
    {
        if (columns.Count == 0)
        {
            return;
        }
        var locked = workbook.CreateCellStyle();
        locked.IsLocked = true;
        var unlocked = workbook.CreateCellStyle();
        unlocked.IsLocked = false;

        for (int c = 0; c < columns.Count; c++)
        {
            sheet.SetDefaultColumnStyle(c, columns[c].IsLocked ? locked : unlocked);
        }
    }

    // ---------- 数据 ----------

    private void WriteData(
        ISheet sheet,
        IWorkbook workbook,
        IReadOnlyList<ExcelColumn> columns,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        IFormulaResolver formulaResolver,
        IEnumerable<object> rows,
        ExcelExportOptions options,
        WriterStyles styles)
    {
        int colCount = columns.Count;

        // A1：公式预检——仅当存在公式列时才为每行构建地址映射
        bool hasFormula = false;
        bool[] formulaCols = new bool[colCount];
        for (int c = 0; c < colCount; c++)
        {
            if (!string.IsNullOrWhiteSpace(columns[c].Formula))
            {
                hasFormula = true;
                formulaCols[c] = true;
            }
        }

        // A2：列字母预计算
        string[] columnAddresses = new string[colCount];
        for (int c = 0; c < colCount; c++)
        {
            columnAddresses[c] = GetColumnAddress(c);
        }

        // A7：表头字节长度 / 当前列宽 预计算（避免每格重复计算）
        int[] headerBytes = new int[colCount];
        double[] widthUnits = new double[colCount];
        int[] maxContentBytes = new int[colCount];
        for (int c = 0; c < colCount; c++)
        {
            headerBytes[c] = Encoding.UTF8.GetBytes(columns[c].Name).Length + 2;
            widthUnits[c] = sheet.GetColumnWidth(c);
        }

        int total = options.KnownRowCount
                    ?? (rows as ICollection<object>)?.Count
                    ?? (rows as IReadOnlyCollection<object>)?.Count ?? 0;
        var progress = options.Progress;

        // A8：导出路径按列预解析 getter（默认访问器时），避免每格 GetOrAdd；
        // 行类型从首个元素解析，遇异构集合时重新解析。
        bool defaultAccessor = accessor is Accessors.ReflectionEntityAccessor;
        Func<object, object?>?[]? getters = null;
        Type? rowType = null;

        int r = 0;
        foreach (var row in rows)
        {
            if (rowType != row.GetType())
            {
                rowType = row.GetType();
                if (defaultAccessor && rowType != null)
                {
                    getters = new Func<object, object?>[colCount];
                    for (int c = 0; c < colCount; c++)
                    {
                        getters[c] = Accessors.ReflectionEntityAccessor.GetGetter(rowType, columns[c].Field);
                    }
                }
                else
                {
                    getters = null;
                }
            }

            int dataRowIndex = HeaderRowCount + r;
            int excelRowNumber = dataRowIndex + 1; // Excel 行号（1 基）
            IRow dataRow = sheet.CreateRow(dataRowIndex);

            // 公式占位符映射：列 Field -> 当前行单元格地址（仅公式列存在时构建）
            Dictionary<string, string>? addressMap = null;
            if (hasFormula)
            {
                addressMap = new Dictionary<string, string>(colCount);
                for (int c = 0; c < colCount; c++)
                {
                    addressMap[columns[c].Field] = columnAddresses[c] + excelRowNumber;
                }
            }

            bool alternate = options.AlternateRowFillColor > 0 && r % 2 == 1;
            for (int c = 0; c < colCount; c++)
            {
                var col = columns[c];
                ICell cell = dataRow.CreateCell(c, CellType.String);
                object? value = getters == null ? accessor.Read(row, col.Field) : getters[c]?.Invoke(row);

                if (formulaCols[c])
                {
                    cell.SetCellFormula(formulaResolver.Resolve(col.Formula, addressMap!));
                }
                else if (IsZeroWhite(value, col))
                {
                    cell.SetCellValue("");
                }
                else
                {
                    WriteCellValue(cell, value, col, converter);
                }

                cell.CellStyle = styles.CellStyleFor(value, col, alternate, options.AlternateRowFillColor, options.EnableBorders);
                AutoWidth(sheet, c, col, options, value, formulaCols[c], headerBytes[c], ref widthUnits[c], ref maxContentBytes[c]);
            }

            if (options.RowHeight > 0)
            {
                dataRow.HeightInPoints = (float)options.RowHeight;
            }

            r++;
            progress?.Report(total > 0 ? (double)r / total : 1.0);
        }

        // 公式结果处理：默认由 NPOI 预计算并写入缓存值（保证导出后直接可见）；
        // 也可跳过求值，改为标记"打开时全量重算"（calcPr.fullCalcOnLoad），
        // 由 Excel 打开时自动计算，导出显著提速。
        if (hasFormula)
        {
            if (options.CalculateFormulasOnExport)
            {
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            }
            else if (workbook is XSSFWorkbook xssfWorkbook)
            {
                // 注意：NPOI 2.7.2 的 SetForceFormulaRecalculation 不能可靠写出 fullCalcOnLoad，
                // 这里直接操作 CT_CalcPr（与 SheetProtector 的写法一致）
                var calcPr = xssfWorkbook.GetCTWorkbook().calcPr ?? xssfWorkbook.GetCTWorkbook().AddNewCalcPr();
                calcPr.fullCalcOnLoad = true;
            }
        }
    }

    private static void WriteCellValue(ICell cell, object? value, ExcelColumn col, ICellValueConverter converter)
    {
        if (value == null)
        {
            cell.SetCellValue("");
            return;
        }
        // 值映射：实体代码值 -> 显示文本（如 0 -> 女 / 1 -> 男）
        if (ValueMapper.TryGetDisplayText(col, value, out string? displayText))
        {
            cell.SetCellValue(displayText);
            return;
        }
        if (col.WriteAsNumeric && converter.IsNumeric(value))
        {
            cell.SetCellValue(Convert.ToDouble(value));
            return;
        }
        if (value is bool b)
        {
            cell.SetCellValue(b);
            return;
        }
        if (value is DateTime dt)
        {
            cell.SetCellValue(dt);
            return;
        }
        cell.SetCellValue(converter.ToCellText(value) ?? "");
    }

    private static bool IsZeroWhite(object? value, ExcelColumn col)
    {
        if (!col.ZeroShowWhiteSpace || value == null)
        {
            return false;
        }
        return value.ToString() == "0" || value.ToString() == "0.0";
    }

    // ---------- 列宽 / 隐藏 ----------

    private static void AutoWidth(
        ISheet sheet,
        int columnIndex,
        ExcelColumn col,
        ExcelExportOptions options,
        object? value,
        bool isFormula,
        int headerBytes,
        ref double widthUnits,
        ref int maxContentBytes)
    {
        if (col.Width is > 0)
        {
            double w = col.Width.Value * 256;
            if (w > widthUnits)
            {
                sheet.SetColumnWidth(columnIndex, (int)w);
                widthUnits = w;
            }
            return;
        }

        // 内容宽度：用原始值文本（避免对每个数值格调用 DataFormatter / cell.ToString）；
        // 公式列以表头宽度为准（公式结果多为短数值）。日期用显示格式。
        if (!isFormula)
        {
            string text = value switch
            {
                DateTime dt => dt.ToString("yyyy-MM-dd"),
                _ => value?.ToString() ?? "",
            };
            if (text.Length + 2 > maxContentBytes)
            {
                int contentBytes = Encoding.UTF8.GetBytes(text).Length + 2;
                if (contentBytes > maxContentBytes)
                {
                    maxContentBytes = contentBytes;
                }
            }
        }

        int len = Math.Max(headerBytes, maxContentBytes);
        len = Math.Max(len, 8);
        len = Math.Min(len, options.MaxColumnWidth);

        if (len * 256 > widthUnits)
        {
            sheet.SetColumnWidth(columnIndex, len * 256);
            widthUnits = len * 256;
        }
    }

    private static void ApplyColumnVisibility(ISheet sheet, int columnIndex, ExcelColumn col)
    {
        if (!col.Visible)
        {
            sheet.SetColumnHidden(columnIndex, true);
        }
    }

    // ---------- 数据验证 ----------

    private void ApplyValidation(ISheet sheet, IReadOnlyList<ExcelColumn> columns, ExcelExportOptions options)
    {
        if (columns.Count == 0)
        {
            return;
        }
        var helper = sheet.GetDataValidationHelper();
        int firstDataRow = HeaderRowCount;
        int lastDataRow = Math.Max(options.ValidationMaxRow - 1, HeaderRowCount);

        for (int c = 0; c < columns.Count; c++)
        {
            var col = columns[c];
            IDataValidationConstraint? constraint = BuildConstraint(helper, col);
            if (constraint == null)
            {
                continue;
            }

            var addressList = new CellRangeAddressList(firstDataRow, lastDataRow, c, c);
            var validation = helper.CreateValidation(constraint, addressList);
            sheet.AddValidationData(validation);
        }
    }

    private static IDataValidationConstraint? BuildConstraint(IDataValidationHelper helper, ExcelColumn col)
    {
        // 显式下拉列表（DropdownOptions 优先）
        if (col.DropdownOptions != null && col.DropdownOptions.Count > 0)
        {
            return helper.CreateExplicitListConstraint(col.DropdownOptions.ToArray());
        }

        switch (col.ValidationKind)
        {
            case ValidationKind.List:
            case ValidationKind.FormulaList:
                return string.IsNullOrWhiteSpace(col.ValidationFormula1)
                    ? null
                    : helper.CreateFormulaListConstraint(col.ValidationFormula1);

            case ValidationKind.Integer:
                return helper.CreateintConstraint(OperatorType.BETWEEN, col.ValidationFormula1, col.ValidationFormula2);

            case ValidationKind.Decimal:
                return helper.CreateDecimalConstraint(OperatorType.BETWEEN, col.ValidationFormula1, col.ValidationFormula2);

            case ValidationKind.Date:
                return helper.CreateDateConstraint(OperatorType.BETWEEN, col.ValidationFormula1, col.ValidationFormula2, "yyyy-mm-dd");

            case ValidationKind.CustomFormula:
                return string.IsNullOrWhiteSpace(col.ValidationFormula1)
                    ? null
                    : helper.CreateCustomConstraint(col.ValidationFormula1);

            default:
                return null;
        }
    }

    // ---------- 保护 ----------

    private void ApplyProtection(ISheet sheet, IReadOnlyList<ExcelColumn> columns, ExcelExportOptions options)
    {
        bool hasLocked = columns.Any(x => x.IsLocked);
        if (!hasLocked && options.Protection == null)
        {
            return;
        }
        SheetProtector.Apply(sheet, options.Protection ?? new SheetProtectionOptions());
    }

    // ---------- 工具 ----------

    /// <summary>列索引 -> 列字母（A, B, ..., Z, AA, AB...）</summary>
    public static string GetColumnAddress(int columnIndex)
    {
        string address = "";
        int index = columnIndex;
        while (index >= 0)
        {
            address = (char)('A' + index % 26) + address;
            index = index / 26 - 1;
        }
        return address;
    }

    /// <summary>设置列宽并按需隐藏列（供子类表头实现调用）</summary>
    protected static void ApplyColumn(ISheet sheet, int columnIndex, ExcelColumn col, ExcelExportOptions options)
    {
        ApplyColumnVisibility(sheet, columnIndex, col);

        int len;
        if (col.Width is > 0)
        {
            len = col.Width.Value;
        }
        else
        {
            len = Encoding.UTF8.GetBytes(col.Name).Length + 2;
            len = Math.Max(len, 8);
            len = Math.Min(len, options.MaxColumnWidth);
        }
        sheet.SetColumnWidth(columnIndex, len * 256);
    }
}

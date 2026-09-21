using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Readers;

/// <summary>
/// 分组表头读取策略（二级表头）：
/// 第 1 行分组、第 2 行列名，数据从第 3 行开始；
/// 通过合并单元格识别分组，按"列名_分组名"定位列（支持用户调整列顺序）。
/// </summary>
public class GroupHeaderSheetReader : ISheetReader
{
    public IReadOnlyList<object> Read(
        ISheet sheet,
        IReadOnlyList<ExcelColumn> columns,
        Type entityType,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        ExcelImportOptions options)
    {
        var result = new List<object>();
        if (sheet == null)
        {
            return result;
        }

        // 期望列 key：无分组 -> Name；有分组 -> Name_GroupName
        var expectedLookup = new Dictionary<string, ExcelColumn>(StringComparer.OrdinalIgnoreCase);
        foreach (var col in columns)
        {
            if (!col.Visible || string.IsNullOrWhiteSpace(col.Name))
            {
                continue;
            }
            string key = BuildKey(col);
            if (!expectedLookup.ContainsKey(key))
            {
                expectedLookup[key] = col;
            }
        }

        // 解析工作表的二级表头 -> 列索引 -> key
        var parsed = ParseGroupHeader(sheet, columns, converter);

        // 校验列对应关系：缺失 / 多余 / 只读忽略 / 表头重复
        var duplicateKeys = parsed.Values
            .GroupBy(v => v.Trim(), StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();

        var validation = ColumnValidator.Validate(
            parsed.Values.Select(v => v.Trim()),
            columns,
            keyOf: BuildKey,
            isRequired: col => options.IgnoreReadOnlyColumns ? ColumnValidator.IsWritable(entityType, col.Field) : true,
            duplicateKeys: duplicateKeys);
        ColumnValidator.Apply(options, validation);

        // 按列预计算：属性类型数组 + 赋值委托数组（索引 = 实际列号，避免循环内字典查找）
        bool defaultAccessor = accessor is Accessors.ReflectionEntityAccessor;
        bool defaultConverter = converter is Converters.DefaultCellValueConverter;
        int maxCol = parsed.Count == 0 ? 0 : parsed.Keys.Max();
        var propertyTypes = new Type[maxCol + 1];
        Action<object, object?>[]? setters = defaultAccessor ? new Action<object, object?>[maxCol + 1] : null;
        foreach (var kv in parsed)
        {
            if (expectedLookup.TryGetValue(kv.Value.Trim(), out var col))
            {
                propertyTypes[kv.Key] = Accessors.ReflectionEntityAccessor.GetPropertyType(entityType, col.Field);
                if (defaultAccessor)
                {
                    setters![kv.Key] = defaultConverter
                        ? Accessors.ReflectionEntityAccessor.GetSetter(entityType, col.Field)
                        : Accessors.ReflectionEntityAccessor.GetWriter(entityType, col.Field);
                }
            }
        }

        // A4：缓存的行对象工厂
        var factory = Accessors.ReflectionEntityAccessor.GetFactory(entityType);

        // C9：进度（按数据区行数估算）
        int dataStart = options.DataStartRowIndex >= 3 ? options.DataStartRowIndex : 2;
        long total = Math.Max(0, (long)sheet.LastRowNum - dataStart + 1);
        long processed = 0;
        var progress = options.Progress;

        for (int r = dataStart; r <= sheet.LastRowNum; r++)
        {
            IRow? row = sheet.GetRow(r);
            if (row == null)
            {
                continue;
            }
            processed++;

            object entity;
            bool any = false;
            try
            {
                entity = factory();
                foreach (var kv in parsed)
                {
                    if (!expectedLookup.TryGetValue(kv.Value.Trim(), out var col))
                    {
                        continue;
                    }
                    if (!ColumnValidator.IsWritable(entityType, col.Field))
                    {
                        continue; // 只读/不存在属性列：不参与写入
                    }
                    ICell? cell = row.GetCell(kv.Key);
                    var value = converter.ReadCell(cell, propertyTypes[kv.Key]);
                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                    {
                        any = true;
                    }
                    var setter = setters?[kv.Key];
                    if (setter != null)
                    {
                        setter(entity, value);
                    }
                    else
                    {
                        accessor.Write(entity, col.Field, value);
                    }
                }

                // C7：源行号回填
                if (!string.IsNullOrWhiteSpace(options.SourceRowProperty))
                {
                    accessor.Write(entity, options.SourceRowProperty!, r + 1);
                }
            }
            catch (Exception ex)
            {
                // C6：逐行错误收集（开启时不抛整批，记录后跳过该行）
                if (options.CollectRowErrors)
                {
                    options.Issues.Add(ImportIssue.Error("RowWriteFailed", ex.Message, r + 1));
                    options.Errors.Add($"第 {r + 1} 行解析失败：{ex.Message}");
                    continue;
                }
                throw;
            }

            if (!any && options.SkipEmptyRows)
            {
                continue;
            }
            result.Add(entity);

            progress?.Report(total > 0 ? (double)processed / total : 1.0);
        }

        return result;
    }

    private static string BuildKey(ExcelColumn col)
    {
        return string.IsNullOrWhiteSpace(col.GroupName)
            ? col.Name.Trim()
            : col.Name.Trim() + "_" + col.GroupName.Trim();
    }

    /// <summary>
    /// 解析二级表头：返回 列索引 -> "列名_分组名"（无分组为列名）
    /// </summary>
    private static Dictionary<int, string> ParseGroupHeader(ISheet sheet, IReadOnlyList<ExcelColumn> columns, ICellValueConverter converter)
    {
        var columnCells = new Dictionary<int, string>();
        IRow row0 = sheet.GetRow(0);
        IRow row1 = sheet.GetRow(1);
        int cellCount = Math.Max(row0?.LastCellNum ?? 0, row1?.LastCellNum ?? 0);

        // 构建分组定义（连续的同组列合并为一个分组）
        var groups = new List<(string Name, string Text, int Count)>();
        int i = 0;
        while (i < columns.Count)
        {
            var col = columns[i];
            if (string.IsNullOrWhiteSpace(col.GroupName))
            {
                groups.Add(("", "", 1));
                i++;
            }
            else
            {
                int s = i;
                int e = i;
                while (e + 1 < columns.Count && columns[e + 1].GroupName == col.GroupName)
                {
                    e++;
                }
                groups.Add((col.GroupName, col.GroupText, e - s + 1));
                i = e + 1;
            }
        }

        int groupLen = 1;
        string groupName = "";
        int startRow = 0;

        for (int c = 0; c < cellCount; c++)
        {
            string headerText = GetCellText(sheet, 0, c, converter);
            var match = groups.FirstOrDefault(g => g.Text == headerText && !string.IsNullOrEmpty(g.Name));

            if (match.Name != null && match.Name != "")
            {
                groupLen = match.Count;
                startRow = 1;
                groupName = match.Name;
                string sub = GetCellText(sheet, 1, c, converter);
                columnCells[c] = sub + "_" + groupName;
                groupLen--;
            }
            else
            {
                string text = startRow == 1 ? GetCellText(sheet, 1, c, converter) : headerText;
                columnCells[c] = startRow == 1 && groupName != "" ? text + "_" + groupName : text;
                groupLen--;
            }

            if (groupLen <= 0)
            {
                startRow = 0;
                groupName = "";
                groupLen = 1;
            }
        }

        return columnCells;
    }

    private static string GetCellText(ISheet sheet, int rowIndex, int columnIndex, ICellValueConverter converter)
    {
        IRow? row = sheet.GetRow(rowIndex);
        if (row == null)
        {
            return "";
        }
        ICell? cell = row.GetCell(columnIndex);
        return converter.ReadCell(cell, typeof(string))?.ToString()?.Trim() ?? "";
    }
}

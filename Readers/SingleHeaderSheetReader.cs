using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Readers;

/// <summary>
/// 单表头读取策略：第 1 行为表头，按表头文字定位列，第 2 行起读取数据
/// </summary>
public class SingleHeaderSheetReader : ISheetReader
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

        int headerRowIndex = Math.Max(0, options.DataStartRowIndex - 1);
        IRow? header = sheet.GetRow(headerRowIndex);
        if (header == null)
        {
            return result;
        }

        // 表头文字 -> 列索引（首个出现位置），并收集重复列名
        var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var duplicateKeys = new List<string>();
        for (int c = 0; c < header.LastCellNum; c++)
        {
            string text = converter.ReadCell(header.GetCell(c), typeof(string))?.ToString()?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }
            if (headerMap.ContainsKey(text))
            {
                duplicateKeys.Add(text);
            }
            else
            {
                headerMap[text] = c;
            }
        }

        // 校验列对应关系：缺失 / 多余 / 只读忽略 / 表头重复
        var validation = ColumnValidator.Validate(
            headerMap.Keys,
            columns,
            keyOf: col => col.Name.Trim(),
            isRequired: col => options.IgnoreReadOnlyColumns ? ColumnValidator.IsWritable(entityType, col.Field) : true,
            duplicateKeys: duplicateKeys);
        ColumnValidator.Apply(options, validation);

        // 列索引 -> 列定义（仅可写列，含缺少列检查）
        var colIndexMap = new Dictionary<int, ExcelColumn>();
        foreach (var col in columns)
        {
            if (!col.Visible || string.IsNullOrWhiteSpace(col.Name))
            {
                continue;
            }
            if (!ColumnValidator.IsWritable(entityType, col.Field))
            {
                continue; // 只读/不存在属性列：不参与写入
            }
            if (headerMap.TryGetValue(col.Name.Trim(), out int idx))
            {
                colIndexMap[idx] = col;
            }
        }

        // 按列预计算：属性类型数组 + 赋值委托数组（索引 = 实际列号，避免循环内字典查找）
        bool defaultAccessor = accessor is Accessors.ReflectionEntityAccessor;
        bool defaultConverter = converter is Converters.DefaultCellValueConverter;
        int maxCol = colIndexMap.Count == 0 ? 0 : colIndexMap.Keys.Max();
        var propertyTypes = new Type[maxCol + 1];
        Action<object, object?>?[]? setters = defaultAccessor ? new Action<object, object?>?[maxCol + 1] : null;
        foreach (var kv in colIndexMap)
        {
            propertyTypes[kv.Key] = Accessors.ReflectionEntityAccessor.GetPropertyType(entityType, kv.Value.Field);
            if (defaultAccessor)
            {
                // 默认转换器：值已是目标类型 -> 用纯赋值委托（省去二次转换）；自定义转换器 -> 用"转换+赋值"委托
                setters![kv.Key] = defaultConverter
                    ? Accessors.ReflectionEntityAccessor.GetSetter(entityType, kv.Value.Field)
                    : Accessors.ReflectionEntityAccessor.GetWriter(entityType, kv.Value.Field);
            }
        }

        // A4：缓存的行对象工厂
        var factory = Accessors.ReflectionEntityAccessor.GetFactory(entityType);

        // C9：进度（按数据区行数估算）
        long total = Math.Max(0, (long)sheet.LastRowNum - options.DataStartRowIndex + 1);
        long processed = 0;
        var progress = options.Progress;

        for (int r = options.DataStartRowIndex; r <= sheet.LastRowNum; r++)
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
                foreach (var kv in colIndexMap)
                {
                    ICell? cell = row.GetCell(kv.Key);
                    object? value;
                    if (kv.Value.ValueMap != null)
                    {
                        // 值映射列：先按原始文本读，再反查为实体代码值（避免先转类型丢失显示文本）
                        value = Converters.ValueMapper.ResolveValue(kv.Value, converter.ReadCell(cell, typeof(string))?.ToString(), propertyTypes[kv.Key]);
                    }
                    else
                    {
                        value = converter.ReadCell(cell, propertyTypes[kv.Key]);
                    }
                    if (IsMeaningful(value))
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
                        accessor.Write(entity, kv.Value.Field, value);
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

    private static bool IsMeaningful(object? value)
    {
        return value != null && !string.IsNullOrEmpty(value.ToString());
    }
}

using System.Reflection;
using MagicFrame.Excel.Exceptions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;

namespace MagicFrame.Excel.Readers;

/// <summary>
/// 列对应关系校验结果：
/// 缺失列、多余列、被忽略列、表头重复列
/// </summary>
public class ColumnCorrespondenceResult
{
    /// <summary>缺失列：实体期望的可见列，但 Excel 中找不到</summary>
    public IReadOnlyList<string> MissingColumns { get; }

    /// <summary>多余列：Excel 表头中存在，但无法对应任何实体列</summary>
    public IReadOnlyList<string> UnexpectedColumns { get; }

    /// <summary>被忽略列：映射到只读/不存在属性，不参与写入，也不视为缺失</summary>
    public IReadOnlyList<string> IgnoredColumns { get; }

    /// <summary>表头重复列：Excel 表头中重复出现的列名</summary>
    public IReadOnlyList<string> DuplicateColumns { get; }

    internal ColumnCorrespondenceResult(
        IReadOnlyList<string> missingColumns,
        IReadOnlyList<string> unexpectedColumns,
        IReadOnlyList<string> ignoredColumns,
        IReadOnlyList<string> duplicateColumns)
    {
        MissingColumns = missingColumns;
        UnexpectedColumns = unexpectedColumns;
        IgnoredColumns = ignoredColumns;
        DuplicateColumns = duplicateColumns;
    }

    public bool HasProblem => MissingColumns.Count > 0 || UnexpectedColumns.Count > 0 || DuplicateColumns.Count > 0;
}

/// <summary>
/// 导入时校验 Excel 列与实体列的对应关系。
/// <para>
/// 校验四类情况：
/// 1) 缺失列：实体期望的可见列在 Excel 中不存在；
/// 2) 多余列：Excel 中的列无法对应任何实体列；
/// 3) 只读列：列映射到只读/不存在的属性，按选项决定忽略或视为缺失；
/// 4) 重复列：Excel 表头出现重复列名，映射存在歧义。
/// </para>
/// </summary>
public static class ColumnValidator
{
    /// <summary>
    /// 校验列对应关系（按各读取策略自身的列 key 约定）
    /// </summary>
    /// <param name="actualKeys">Excel 表头中解析出的列 key（去重、已 Trim）</param>
    /// <param name="columns">期望列定义</param>
    /// <param name="keyOf">列定义 → key（单表头为列名；分组表头为"列名_分组名"）</param>
    /// <param name="isRequired">该列是否视为必需/可写（读返回 false 的列会被忽略）</param>
    /// <param name="duplicateKeys">Excel 表头中重复出现的列 key</param>
    public static ColumnCorrespondenceResult Validate(
        IEnumerable<string> actualKeys,
        IReadOnlyList<ExcelColumn> columns,
        Func<ExcelColumn, string> keyOf,
        Func<ExcelColumn, bool> isRequired,
        IEnumerable<string> duplicateKeys)
    {
        var actual = new HashSet<string>(actualKeys, StringComparer.OrdinalIgnoreCase);

        // 全部已知 key（含隐藏列）：用于判定"多余列"，避免把隐藏列误判为多余
        var knownKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // 可见列：用于判定"缺失/忽略"
        var visibleByKey = new Dictionary<string, ExcelColumn>(StringComparer.OrdinalIgnoreCase);

        foreach (var col in columns)
        {
            string key = keyOf(col);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }
            knownKeys.Add(key);
            if (col.Visible)
            {
                visibleByKey[key] = col;
            }
        }

        var missing = new List<string>();
        var ignored = new List<string>();
        foreach (var kv in visibleByKey)
        {
            if (!isRequired(kv.Value))
            {
                ignored.Add(kv.Value.Name);
            }
            else if (!actual.Contains(kv.Key))
            {
                missing.Add(kv.Value.Name);
            }
        }

        var unexpected = actual.Where(k => !knownKeys.Contains(k)).ToList();
        var duplicates = duplicateKeys.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        return new ColumnCorrespondenceResult(missing, unexpected, ignored, duplicates);
    }

    /// <summary>
    /// 按导入选项落地校验结果：记录到 Errors + Issues，并按选项抛出异常
    /// </summary>
    public static void Apply(ExcelImportOptions options, ColumnCorrespondenceResult result)
    {
        if (result == null)
        {
            return;
        }

        if (result.MissingColumns.Count > 0)
        {
            string msg = "缺少列：" + string.Join(",", result.MissingColumns);
            options.Errors.Add(msg);
            options.Issues.Add(ImportIssue.Error("MissingColumn", msg));
            if (options.ThrowOnMissingColumns)
            {
                throw new ExcelImportException(msg);
            }
        }

        if (result.UnexpectedColumns.Count > 0)
        {
            string msg = "存在无法对应实体列的列：" + string.Join(",", result.UnexpectedColumns);
            options.Errors.Add(msg);
            options.Issues.Add(ImportIssue.Warning("UnexpectedColumn", msg));
            if (options.ThrowOnUnexpectedColumns)
            {
                throw new ExcelImportException("存在无法对应实体列的列：" + string.Join(",", result.UnexpectedColumns));
            }
        }

        if (result.IgnoredColumns.Count > 0)
        {
            string msg = "以下列映射到只读或不存在的属性，已忽略：" + string.Join(",", result.IgnoredColumns);
            options.Errors.Add(msg);
            options.Issues.Add(ImportIssue.Info("IgnoredColumn", msg));
        }

        if (result.DuplicateColumns.Count > 0)
        {
            string msg = "表头存在重复列名，按首个出现位置映射：" + string.Join(",", result.DuplicateColumns);
            options.Errors.Add(msg);
            options.Issues.Add(ImportIssue.Warning("DuplicateColumn", msg));
        }
    }

    /// <summary>
    /// 判断列 Field 是否为可写属性。
    /// <para>
    /// 属性存在但为只读（get-only）时返回 false；属性不存在时返回 true，
    /// 以便自定义访问器（如字典行）可处理任意字段，而不会把列误判为只读。
    /// </para>
    /// </summary>
    public static bool IsWritable(Type entityType, string field)
    {
        if (entityType == null || string.IsNullOrEmpty(field))
        {
            return true;
        }
        var prop = entityType.GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        return prop == null || prop.CanWrite;
    }
}

using System.Text.RegularExpressions;
using MagicFrame.Excel.Abstractions;

namespace MagicFrame.Excel.Formulas;

/// <summary>
/// 默认公式解析器：把 {Field} 占位符替换为列 Field 对应的单元格地址
/// </summary>
public class DefaultFormulaResolver : IFormulaResolver
{
    public static readonly DefaultFormulaResolver Instance = new();

    private static readonly Regex PlaceholderRegex = new(@"\{([^{}]+)\}", RegexOptions.Compiled);

    public string Resolve(string template, IReadOnlyDictionary<string, string> fieldToAddress)
    {
        if (string.IsNullOrEmpty(template))
        {
            return template;
        }
        return PlaceholderRegex.Replace(template, m =>
        {
            string key = m.Groups[1].Value;
            return fieldToAddress.TryGetValue(key, out string? address) ? address : m.Value;
        });
    }
}

namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// 公式占位符解析器：把形如 "IFERROR({A}+{B},\"\")" 的模板，
/// 按"列 Field → 单元格地址（如 B3）"的映射替换为真实公式。
/// 默认实现为花括号正则替换；如需自定义语法实现本接口即可。
/// </summary>
public interface IFormulaResolver
{
    /// <summary>
    /// 解析公式模板
    /// </summary>
    /// <param name="template">公式模板，无需等号开头</param>
    /// <param name="fieldToAddress">列 Field 到当前行单元格地址的映射</param>
    /// <returns>解析后的公式（可直接 SetCellFormula）</returns>
    string Resolve(string template, IReadOnlyDictionary<string, string> fieldToAddress);
}

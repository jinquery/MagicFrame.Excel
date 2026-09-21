using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// 单元格值与实体值之间的转换器。
/// 默认实现处理 数值/文本/布尔/日期/公式求值 等类型；
/// 如需自定义类型（枚举映射、JSON、加密等），实现本接口并在 Builder 中注册。
/// </summary>
public interface ICellValueConverter
{
    /// <summary>
    /// 读取单元格并转换为目标属性类型
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="targetType">目标类型。单元格内容要转换成的类型</param>
    /// <returns></returns>
    object? ReadCell(ICell? cell, Type targetType);

    /// <summary>
    /// 将实体值转换为写入单元格的文本
    /// </summary>
    /// <param name="value">内容</param>
    /// <returns>返回转换后的文本</returns>
    string? ToCellText(object? value);

    /// <summary>
    /// 判断实体值是否为数值（决定是否按数值单元格写入）
    /// </summary>
    /// <param name="value">实体的值</param>
    /// <returns>true表示为数字，false非数字</returns>
    bool IsNumeric(object? value);
}

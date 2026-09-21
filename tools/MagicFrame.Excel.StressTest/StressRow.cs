namespace MagicFrame.Excel.StressTest;

/// <summary>
/// 压力测试实体：30 个属性（10 字符串 / 10 数值 / 5 金额 / 3 日期 / 2 布尔），
/// 覆盖库内常用的所有单元格类型写入与读取路径。
/// </summary>
public class StressRow
{
    // 字符串（10）
    public string P01 { get; set; } = "";
    public string P02 { get; set; } = "";
    public string P03 { get; set; } = "";
    public string P04 { get; set; } = "";
    public string P05 { get; set; } = "";
    public string P06 { get; set; } = "";
    public string P07 { get; set; } = "";
    public string P08 { get; set; } = "";
    public string P09 { get; set; } = "";
    public string P10 { get; set; } = "";

    // 数值 double（10）
    public double P11 { get; set; }
    public double P12 { get; set; }
    public double P13 { get; set; }
    public double P14 { get; set; }
    public double P15 { get; set; }
    public double P16 { get; set; }
    public double P17 { get; set; }
    public double P18 { get; set; }
    public double P19 { get; set; }
    public double P20 { get; set; }

    // 金额 decimal（5）
    public decimal P21 { get; set; }
    public decimal P22 { get; set; }
    public decimal P23 { get; set; }
    public decimal P24 { get; set; }
    public decimal P25 { get; set; }

    // 日期 DateTime（3）
    public DateTime P26 { get; set; }
    public DateTime P27 { get; set; }
    public DateTime P28 { get; set; }

    // 布尔 bool（2）
    public bool P29 { get; set; }
    public bool P30 { get; set; }

    /// <summary>公式列（压测公式求值用，正常模式不参与）</summary>
    public decimal Total { get; set; }

    /// <summary>按行号生成一行确定性数据</summary>
    public static StressRow Create(int i)
    {
        var baseDate = new DateTime(2020, 1, 1);
        return new StressRow
        {
            P01 = $"姓名{i:D4}",
            P02 = $"部门{i % 8}",
            P03 = $"职位-{i % 5}",
            P04 = $"编号{i:D6}",
            P05 = $"备注{i % 3}",
            P06 = $"A{i:D3}",
            P07 = $"B{i:D3}",
            P08 = $"C{i:D3}",
            P09 = $"D{i:D3}",
            P10 = $"E{i:D3}",

            P11 = i * 1.1,
            P12 = i * 2.2,
            P13 = i * 3.3,
            P14 = i * 4.4,
            P15 = i * 5.5,
            P16 = i * 6.6,
            P17 = i * 7.7,
            P18 = i * 8.8,
            P19 = i * 9.9,
            P20 = i * 10.1,

            P21 = i * 1.25m,
            P22 = i * 2.5m,
            P23 = i * 3.75m,
            P24 = i * 4.5m,
            P25 = i * 5.5m,

            P26 = baseDate.AddDays(i),
            P27 = baseDate.AddDays(i + 1),
            P28 = baseDate.AddDays(i + 2),

            P29 = i % 2 == 0,
            P30 = i % 3 == 0,
        };
    }
}

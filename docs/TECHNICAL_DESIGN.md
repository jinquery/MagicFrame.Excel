# MagicFrame.Excel 技术设计文档

> 版本：1.0.0
> 适用代码：`MagicFrame.Excel` 类库（net6.0 / NPOI 2.7.2）
> 一句话定位：**基于 NPOI 的通用 Excel 生成/解析类库**，以「泛型集合 ⇄ Excel」为核心，内置公式、整列锁定、数据验证下拉、多级（分组）表头；整体遵循**对修改关闭、对扩展开放**原则。

---

## 目录

1. [设计目标与原则](#1-设计目标与原则)
2. [技术栈与环境](#2-技术栈与环境)
3. [整体架构](#3-整体架构)
4. [核心概念模型](#4-核心概念模型)
5. [使用说明](#5-使用说明)
6. [扩展说明](#6-扩展说明)
7. [扩展的注意事项](#7-扩展的注意事项)
8. [扩展的边界](#8-扩展的边界)
9. [构建与测试](#9-构建与测试)

---

## 1. 设计目标与原则

### 1.1 设计目标

| 目标 | 说明 |
| --- | --- |
| 泛型化 | 以 `IEnumerable<T>` 作为输入/输出，任意 POCO 均可直接导出/解析 |
| 声明式列配置 | 通过 `[ExcelColumn]` 特性声明列：列名、顺序、公式、锁定、下拉、分组、宽度等 |
| 整列锁定 | 锁定指定列，同时放开「新增行 / 删除行 / 调整格式」，贴近真实业务模板 |
| 公式支持 | 占位符式公式模板（`{Field}`），按行替换为真实单元格地址，导出时预计算 |
| 多级表头 | 内置「单表头」「分组表头」两种布局，可扩展更多布局 |
| 开闭原则 | 核心引擎固定流水线（对修改关闭），所有可定制点通过接口 + Builder 注入（对扩展开放） |

### 1.2 设计原则

1. **单一事实来源**：`ExcelColumn` 是「列定义」的唯一数据模型。特性发现、程序化构建都收敛到它，生成与解析共用同一份定义。
2. **策略可插拔**：列元数据、实体读写、值转换、公式解析、表头布局全部抽象为接口，通过 `ExcelEngineBuilder` 注册/替换。
3. **无跨调用状态**：写入器内部按「每个 Workbook 独立创建样式」（`WriterStyles`），因此各策略实现可安全地作为单例复用。
4. **纯 POCO**：实体无需继承任何基类、无需实现任何接口，仅需可被 `new()` 实例化（满足 `where T : new()`）。

---

## 2. 技术栈与环境

| 项 | 值 |
| --- | --- |
| 目标框架 | `net6.0` |
| 语言版本 | `C# 10`（`ImplicitUsings`、`Nullable` 开启） |
| 依赖 | `NPOI 2.7.2` |
| 文件格式 | **仅 `.xlsx`（XSSF / OpenXML）**，引擎内部固定使用 `XSSFWorkbook` |
| 测试 | `xUnit`（`tests/MagicFrame.Excel.Tests`，70 个用例） |
| 演示 | `demos/MagicFrame.Excel.Demo`（控制台） |

---

## 3. 整体架构

### 3.1 目录结构

```
MagicFrame.Excel/
├── Abstractions/          # 扩展点接口（IColumnProvider / IEntityAccessor / ICellValueConverter /
│                          #   IFormulaResolver / ISheetWriter / ISheetReader）
├── Accessors/             # 默认实体访问器（反射 + 表达式编译，含写入快路径委托）
├── Attributes/            # [ExcelColumn] 特性
├── Converters/            # 默认值转换器 + 类型转换工具（公式求值器按工作簿缓存）
├── Engine/                # 引擎门面、构建器、导出结果、多 Sheet 规格
├── Exceptions/            # ExcelImportException
├── Formulas/              # 默认公式解析器
├── Model/                 # ExcelColumn / CellAlignment / HeaderKinds / ValidationKind
├── Options/               # 导出/导入/保护 选项 + ImportIssue（结构化问题）
├── Protection/            # 工作表保护器（CT_SheetProtection）
├── Providers/             # 默认列提供器（特性发现 + 按类型缓存）
├── Readers/               # 读取策略（单/分组表头）+ 注册表 + 列对应校验 + 快路径写入
├── Writers/               # 写入策略（单/分组表头）+ 注册表 + 样式集合
├── tests/MagicFrame.Excel.Tests/   # xUnit 测试（70 个用例）
├── demos/MagicFrame.Excel.Demo/     # 控制台演示
├── tools/MagicFrame.Excel.StressTest/  # 独立压力测试（5000×30 计时、公式对比）
├── docs/                  # 技术设计文档 + drawio 图表
└── .github/workflows/     # CI（build + test + 压测 + tag 触发 pack）
```

### 3.2 核心流程（固定流水线）

**导出（Export）**

```
ResolveColumns(T) ──► 新建 XSSFWorkbook ──► 按 HeaderKind 取写入策略 ──► Write(sheet, ...)
   列提供器或选项                    │                                    │
                                     │                              ├─ 铺默认列样式（整列锁定）
                                     │                              ├─ 写表头
                                     │                              ├─ 写数据（公式替换/求值）
                                     │                              ├─ 下拉验证
                                     │                              ├─ 冻结表头
                                     │                              └─ 工作表保护
                                     ▼
                         ToResult：Workbook / Bytes / Base64 / Stream
```

**导入（Import）**

```
打开 XSSFWorkbook ──► 按 SheetName 或 SheetIndex 选 Sheet ──► 按 HeaderKind 取读取策略
        ──► Read(sheet, columns, ...) 按列名定位 → 逐行构造实体 → IList<T>
```

### 3.3 扩展点总览

| 扩展点 | 接口 | 默认实现 | 职责 |
| --- | --- | --- | --- |
| 列元数据来源 | `IColumnProvider` | `AttributeColumnProvider` | 由实体类型产出 `IReadOnlyList<ExcelColumn>` |
| 实体读写 | `IEntityAccessor` | `ReflectionEntityAccessor` | 按 `Field` 读写实体属性 |
| 单元格值转换 | `ICellValueConverter` | `DefaultCellValueConverter` | 单元格 ⇄ 实体值（含公式求值） |
| 公式占位符 | `IFormulaResolver` | `DefaultFormulaResolver` | 将 `{Field}` 替换为单元格地址 |
| 表头写入 | `ISheetWriter` | `SingleHeaderSheetWriter` / `GroupHeaderSheetWriter` | 按布局写入一个 Sheet |
| 表头读取 | `ISheetReader` | `SingleHeaderSheetReader` / `GroupHeaderSheetReader` | 按布局解析一个 Sheet |

---

## 4. 核心概念模型

### 4.1 ExcelColumn（列定义，单一事实来源）

| 成员 | 类型 | 说明 |
| --- | --- | --- |
| `Name` | string | 表头显示文字（导入时按此定位列） |
| `Field` | string | 实体属性名 |
| `Order` | int | 列顺序 |
| `Visible` | bool | 是否显示；`false` 为隐藏列（保留位置） |
| `IsLocked` | bool | 整列锁定 |
| `Formula` | string | 公式模板，如 `IFERROR({Qty}*{Price},"")`，无需等号 |
| `ZeroShowWhiteSpace` | bool | 0 值显示为空白 |
| `HeaderColor` | short | 表头字体颜色（HSSFColor 索引） |
| `GroupName` / `GroupText` | string | 分组表头：分组标识 / 分组显示文字 |
| `DropdownOptions` | IList\<string\> | 数据验证下拉选项 |
| `ValidationKind` | ValidationKind | 数据验证类型：List / Integer / Decimal / Date / CustomFormula / FormulaList |
| `ValidationFormula1` / `ValidationFormula2` | string | 数据验证参数（区间边界 / 自定义公式 / 列表来源） |
| `Width` | int? | 列宽（字符数）；`null` 按表头文字自适应 |
| `Alignment` | CellAlignment | 水平对齐 |
| `WriteAsNumeric` | bool | 是否按数值写入（false 统一按文本） |
| `NumberFormat` | string | 数字格式掩码，如 `"0.00%"` / `"#,##0.00"` |
| `ValueMap` | IDictionary\<object,string\> | 值映射：实体代码值 → 显示文本（导出替换、导入反查）；`"unknown"` 哨兵键指定异常值兜底代码 |
| `Clone()` | ExcelColumn | 深拷贝（避免引用共享） |

### 4.2 ExcelColumnAttribute（特性版）

与 `ExcelColumn` 一一对应，用于在实体属性上声明。注意：

- `Order` 为 0 时仍按 `Order` 再按 `Name` 排序（`AttributeColumnProvider` 的排序逻辑）。
- `Width` 为 `int`，**0 表示自适应**（特性参数不允许可空值类型，故不使用 `int?`）。
- `DropdownOptions` 为 `string[]`；`ValidationFormula1/2` 为 string；`ValidationKind` / `NumberFormat` / `ValueMappings`（如 `"0:女,1:男,unknown:2"`）可直接在特性中声明。
- 特性继承自 `Attribute`，可设 `Field` 显式指定实体属性名（默认取属性名）。

### 4.3 HeaderKinds（表头类型常量）

| 常量 | 值 | 写入策略 | 读取策略 | HeaderRowCount |
| --- | --- | --- | --- | --- |
| `HeaderKinds.Single` | `"Single"` | 单表头：第 1 行列名 | 按列名定位 | 1 |
| `HeaderKinds.Group` | `"Group"` | 二级表头：第 1 行分组、第 2 行列名 | 按「列名_分组名」定位 | 2 |

### 4.4 选项类

- **ExcelExportOptions**：`SheetName`、`HeaderKind`、`Columns`（程序化列定义，非空时优先于特性发现）、`Protection`、`FreezeHeader`、`FreezeColumns`、`AutoFilter`、`RowHeight`、`AlternateRowFillColor`、`EnableBorders`、`Progress`、`ValidationMaxRow`（下拉作用最大行号，默认 1000）、`MaxColumnWidth`、`CalculateFormulasOnExport`（公式导出策略）。
- **ExcelImportOptions**：`SheetName`（优先）、`SheetIndex`、`HeaderKind`、`DataStartRowIndex`（表头行号，1 基）、`SkipEmptyRows`、`ThrowOnMissingColumns`、`ThrowOnUnexpectedColumns`、`IgnoreReadOnlyColumns`、`CollectRowErrors`、`SourceRowProperty`、`Progress`、`Columns`、`Errors`、`Issues`。
- **SheetProtectionOptions**：`Enabled`、`Password`（默认 `"1234"`）、`ProtectWithoutPassword`、`AllowInsertRows` / `AllowDeleteRows` / `AllowFormat` / `AllowSort` / `AllowAutoFilter`（默认均放开）。
- **ImportIssue**：结构化导入问题（`Severity` / `Code` / `Message` / `RowNumber`），程序化消费优于纯文本 `Errors`。

### 4.5 门面与辅助

- **ExcelEngine**：统一入口，静态 `CreateDefault()` 或由 `ExcelEngineBuilder` 构建。导出 `Export<T>` 系列、`ExportSheets`；导入 `Import<T>` 系列 + `ImportAll<T>`（一键导入全部 Sheet）。
- **ExcelEngineBuilder**：扩展注册入口。
- **ExcelExportResult**：同时承载 `Workbook` / `Bytes` / `Base64` / `Stream`，`IDisposable`。
- **SheetSpec**：多 Sheet 导出的单个 Sheet 规格（`Name` + `RowType` + `Rows` + `Options`）。

---

## 5. 使用说明

### 5.1 快速开始

```csharp
using MagicFrame.Excel.Attributes;
using MagicFrame.Excel.Engine;

public class Employee
{
    [ExcelColumn("姓名", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("年龄", Order = 2)]
    public int Age { get; set; }
}

// 导出
var engine = ExcelEngine.CreateDefault();
var rows = new List<Employee> { new() { Name = "张三", Age = 28 } };
byte[] bytes = engine.ExportBytes(rows, new ExcelExportOptions { SheetName = "员工表" });

// 导入
var back = engine.Import<Employee>(bytes, new ExcelImportOptions { SheetName = "员工表" });
```

### 5.2 导出形态

| 方法 | 返回 | 适用场景 |
| --- | --- | --- |
| `Export<T>` | `ExcelExportResult` | 需要同时拿到工作簿/字节/流 |
| `ExportBytes<T>` | `byte[]` | 直接写文件、传存储 |
| `ExportBase64<T>` | `string` | 前端下载（Data URI） |
| `ExportStream<T>` | `MemoryStream` | 流式返回（Position 已归零） |
| `ExportToFile<T>` | void | 写本地文件 |
| `ExportSheets(IEnumerable<SheetSpec>)` | `ExcelExportResult` | 多 Sheet、可异类型 |
| `ExportSheets<T>(IDictionary<string, IList<T>>)` | `ExcelExportResult` | 多 Sheet、同类型便捷重载 |

### 5.3 导入形态

| 方法 | 输入 |
| --- | --- |
| `Import<T>(byte[])` | 字节 |
| `Import<T>(string base64)` | Base64 |
| `ImportFile<T>(string path)` | 本地文件 |
| `Import<T>(Stream)` | 任意可读流 |

> 导入按 `ExcelImportOptions.SheetName` 选 Sheet，为空时用 `SheetIndex`。`DataStartRowIndex` 指表头所在行（1 基，单表头默认 1、分组表头默认 2），数据从下一行开始。

### 5.4 导入列对应校验

导入时会校验 **Excel 的列与实体列是否对应**，校验结果写入 `ExcelImportOptions.Errors`，并按选项决定是否抛异常：

| 校验项 | 触发条件 | 默认行为 |
| --- | --- | --- |
| 缺失列 | 实体期望的可见列在 Excel 中找不到 | `ThrowOnMissingColumns=true`（默认）抛 `ExcelImportException` |
| 多余列 | Excel 中的列无法对应任何实体列 | 仅写入 `Errors`；`ThrowOnUnexpectedColumns=true` 时抛异常 |
| 只读列忽略 | 列映射到只读（get-only）或不存在属性的实体列 | `IgnoreReadOnlyColumns=true`（默认）不写、不视为缺失，仅记录 `Errors` |
| 表头重复列 | Excel 表头出现重复列名 | 按首个出现位置映射，记录 `Errors` 提示 |

```csharp
var options = new ExcelImportOptions
{
    SheetName = "员工表",
    ThrowOnUnexpectedColumns = true,   // 多余列也视为错误
    // IgnoreReadOnlyColumns = false,  // 需要只读列必须存在时打开
};
try
{
    var rows = engine.Import<Employee>(bytes, options);
}
catch (ExcelImportException ex)
{
    // ex.Message 包含具体缺失/多余列名
}
// 未抛异常时也可读取 options.Errors 查看全部校验提示
foreach (var msg in options.Errors) Console.WriteLine(msg);
```

> **只读列的判定**：仅 `get-only`（`CanWrite == false`）属性视为只读忽略；`private set` 属性经反射仍可写，不算只读。列 `Field` 指向不存在的属性时按可写处理（为自定义访问器如字典行保留通道），写入由访问器自行处理。

### 5.5 特性配置详解

```csharp
public class Product
{
    // 基础列：列名 + 顺序 + 宽度（0=自适应）
    [ExcelColumn("名称", Order = 1, Width = 16)]
    public string Name { get; set; } = "";

    // 下拉验证：限定可选值，作用到 ValidationMaxRow
    [ExcelColumn("分类", Order = 2, DropdownOptions = new[] { "电子", "食品", "服饰" })]
    public string Category { get; set; } = "";

    // 整列锁定：该列（含用户新增行的该列）不可编辑
    [ExcelColumn("工号", Order = 3, IsLocked = true)]
    public string No { get; set; } = "";

    // 公式：{Field} 会被替换为当前行对应单元格地址（如 B2*C2）
    [ExcelColumn("合计", Order = 4, IsLocked = true, Formula = "{Qty}*{Price}")]
    public decimal Total { get; set; }

    // 0 值显示为空白
    [ExcelColumn("销量", Order = 5, ZeroShowWhiteSpace = true)]
    public int Sales { get; set; }

    // 隐藏列：保留位置但隐藏
    [ExcelColumn("内部备注", Order = 6, Visible = false)]
    public string? Remark { get; set; }

    // 分组表头：同一 GroupName 的多列在第 1 行合并显示 GroupText
    [ExcelColumn("业绩", Order = 7, GroupName = "Kpi", GroupText = "考核信息")]
    public int Performance { get; set; }
}
```

**导出选项示例（含锁定保护、分组表头）**：

```csharp
engine.ExportToFile(rows, "out.xlsx", new ExcelExportOptions
{
    SheetName = "考核表",
    HeaderKind = HeaderKinds.Group,               // 分组表头
    Protection = new SheetProtectionOptions       // 锁定 + 放开新增/删除/格式
    {
        Password = "1234",
        AllowInsertRows = true,
        AllowDeleteRows = true,
        AllowFormat = true,
    },
    // CalculateFormulasOnExport = false,          // 大数据量公式导出提速：交给 Excel 打开时重算
});
```

> **公式导出提速**：导出含公式的文件时默认会由 NPOI 预计算（`EvaluateAllFormulaCells`，公式多时较慢）。大数据量可设 `CalculateFormulasOnExport = false`，改为写 `fullCalcOnLoad` 标记，由 Excel 打开时自动重算（详见 7.6）。

### 5.6 程序化列定义（不走特性）

当实体不可加特性（第三方类型、匿名/动态结构、字典行）时，可手工构建 `ExcelColumn`：

```csharp
var columns = new[]
{
    new ExcelColumn { Name = "编号", Field = "Id", Order = 1 },
    new ExcelColumn { Name = "金额", Field = "Amount", Order = 2, Width = 12 },
};

engine.ExportToFile(rows, "out.xlsx",
    new ExcelExportOptions { Columns = columns, SheetName = "表" });
```

> 只要 `ExcelExportOptions.Columns` 或 `ExcelImportOptions.Columns` 非空，就**优先使用**程序化列定义，不再走特性发现。

### 5.7 多 Sheet 导出

```csharp
var specs = new[]
{
    SheetSpec.Of("员工表", employees),
    SheetSpec.Of("考核表", assessments),          // 可异类型
};
byte[] bytes = engine.ExportSheets(specs).Bytes;

// 解析各 Sheet 时用 SheetName 指定
var emp = engine.Import<Employee>(bytes, new ExcelImportOptions { SheetName = "员工表" });

// 或一键导入全部 Sheet（要求所有 Sheet 结构一致，均为 T）
IDictionary<string, IList<Employee>> all = engine.ImportAll<Employee>(bytes);
```

### 5.8 增强功能速览

| 功能 | 入口 | 说明 |
| --- | --- | --- |
| 自动筛选 | `ExcelExportOptions.AutoFilter` | 在表头行添加 `<autoFilter>` |
| 冻结列 | `ExcelExportOptions.FreezeColumns` | 冻结前 N 列，可与 FreezeHeader 叠加 |
| 数字格式 | `ExcelColumn.NumberFormat` | 如 `"0.00%"`、`"#,##0.00"` |
| 数据验证增强 | `ExcelColumn.ValidationKind` + `ValidationFormula1/2` | 整数/小数/日期区间、自定义公式、公式列表 |
| **值映射** | `ExcelColumn.ValueMap`（`"unknown"` 哨兵键 / 特性 `ValueMappings`） | 实体代码值 ⇄ 显示文本（如 0→女、1→男），导入未映射文本回退异常值 |
| 空密码保护 | `SheetProtectionOptions.ProtectWithoutPassword` | 保护但不设密码（Password 置空 + 该开关） |
| 逐行错误收集 | `ExcelImportOptions.CollectRowErrors` | 某行解析失败记录到 `Issues` 并跳过，不整批抛 |
| 源行号回填 | `ExcelImportOptions.SourceRowProperty` | 把 Excel 行号写回实体属性 |
| 进度回调 | `Progress`（导出/导入选项） | `IProgress<double>` 0~1 |
| 样式扩展 | `RowHeight` / `AlternateRowFillColor` / `EnableBorders` | 行高 / 交替行色 / 边框 |
| 结构化问题 | `ExcelImportOptions.Issues` | `ImportIssue`（Severity/Code/Message/Row） |
| 一键多 Sheet 导入 | `ExcelEngine.ImportAll<T>` | 返回 `Dictionary<sheetName, IList<T>>` |

### 5.9 值映射（代码 ⇄ 显示文本）

场景：实体存 0/1，Excel 显示 女/男；导入时 男 → 1、女 → 0；未映射文本回退到"异常值"。

**程序化**：
```csharp
new ExcelColumn
{
    Name = "性别",
    Field = "Gender",
    ValueMap = new Dictionary<object, string>
    {
        [0] = "女",
        [1] = "男",
        [ValueMapper.UnknownKey] = "2",   // "unknown" 哨兵键：导入未映射文本时字段写 2
    },
}
```

**特性**：
```csharp
// "unknown:2" 指定异常值兜底代码；若不写，自动补 unknown:-9999999
[ExcelColumn("性别", Order = 2, ValueMappings = "0:女,1:男,unknown:2")]
public int Gender { get; set; }
```

**行为**：
- 导出：实体值 1 → 单元格写 `男`；0 → `女`；不在映射中的值按原值写入。
- 导入：`男` → 1、`女` → 0；未映射的**非空**文本 → 异常值兜底（取 `unknown` 哨兵键代码，**未提供则默认 `-9999999`**）；空单元格按目标默认值（不视为异常）。
- 特性 `ValueMappings` 中纯数字键按整数解析；也支持单元格里直接填代码值（如数字 `0`）时按原代码还原。
- 程序化 `ValueMap` 未含 `unknown` 键时，兜底同样默认为 `-9999999`（`ValueMapper.DefaultUnknownValue`）。

---

## 6. 扩展说明

扩展的入口只有一个：`ExcelEngineBuilder`。**核心引擎代码零改动**，通过替换/追加扩展实现即可获得新能力。

```csharp
var engine = ExcelEngineBuilder.Create()
    .UseColumnProvider(...)     // 替换列提供器
    .UseEntityAccessor(...)     // 替换实体访问器
    .UseCellValueConverter(...) // 替换值转换器
    .UseFormulaResolver(...)    // 替换公式解析器
    .RegisterSheetWriter(kind, ...)  // 追加/覆盖写入策略
    .RegisterSheetReader(kind, ...)  // 追加/覆盖读取策略
    .Build();
```

### 6.1 IColumnProvider —— 扩展列定义来源

```csharp
public interface IColumnProvider
{
    IReadOnlyList<ExcelColumn> GetColumns(Type entityType);
}
```

**适用**：列定义来自配置文件、数据库、前端下发、动态表头等。

```csharp
public class ConfigColumnProvider : IColumnProvider
{
    private readonly IReadOnlyList<ExcelColumn> _columns;
    public ConfigColumnProvider(IReadOnlyList<ExcelColumn> c) => _columns = c;
    public IReadOnlyList<ExcelColumn> GetColumns(Type entityType) => _columns;
}

var engine = ExcelEngineBuilder.Create().UseColumnProvider(new ConfigColumnProvider(cols)).Build();
```

### 6.2 IEntityAccessor —— 扩展实体读写

```csharp
public interface IEntityAccessor
{
    object? Read(object entity, string field);
    void Write(object entity, string field, object? value);
}
```

**适用**：支持字典行、索引器实体、DataRow、非 POCO 结构。

```csharp
public class DictionaryAccessor : IEntityAccessor
{
    public object? Read(object e, string f)
        => e is IDictionary<string, object?> d && d.TryGetValue(f, out var v) ? v : null;

    public void Write(object e, string f, object? v)
    {
        if (e is IDictionary<string, object?> d) d[f] = v;
    }
}
```

### 6.3 ICellValueConverter —— 扩展值转换

```csharp
public interface ICellValueConverter
{
    object? ReadCell(ICell? cell, Type targetType);  // 读单元格 → 目标属性类型
    string? ToCellText(object? value);               // 实体值 → 单元格文本
    bool IsNumeric(object? value);                   // 是否按数值单元格写入
}
```

**适用**：枚举映射、加密/脱敏、自定义格式化（货币、百分比）、JSON 等。

```csharp
public class CurrencyConverter : ICellValueConverter
{
    // 组合默认实现再改写，推荐做法
    public object? ReadCell(ICell? cell, Type t)
        => DefaultCellValueConverter.Instance.ReadCell(cell, t);

    public string? ToCellText(object? v)
        => v is decimal d ? "¥" + d.ToString("N2") : v?.ToString();

    public bool IsNumeric(object? v) => false; // 一律按文本写入
}
```

### 6.4 IFormulaResolver —— 扩展公式语法

```csharp
public interface IFormulaResolver
{
    string Resolve(string template, IReadOnlyDictionary<string, string> fieldToAddress);
}
```

**适用**：更换占位符语法（如 `[[Field]]`、`{{Field}}`）、接入安全校验、公式白名单。

```csharp
// 把 [[Qty]]*[[Price]] 替换为 A2*B2
public class DoubleBracketResolver : IFormulaResolver { /* 正则替换 */ }
```

### 6.5 ISheetWriter / ISheetReader —— 扩展表头布局（最强大的扩展点）

```csharp
public interface ISheetWriter
{
    int HeaderRowCount { get; }
    void Write(ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns,
        IEntityAccessor accessor, ICellValueConverter converter, IFormulaResolver formulaResolver,
        IEnumerable<object> rows, ExcelExportOptions options);
}

public interface ISheetReader
{
    IReadOnlyList<object> Read(ISheet sheet, IReadOnlyList<ExcelColumn> columns,
        Type entityType, IEntityAccessor accessor, ICellValueConverter converter,
        ExcelImportOptions options);
}
```

**适用**：固定标题 + 列名、多级表头、自定义样式表头、表头行内嵌汇总等。

```csharp
const string FixedTitle = "FixedTitle";
var engine = ExcelEngineBuilder.Create()
    .RegisterSheetWriter(FixedTitle, new FixedTitleWriter())   // 第 1 行固定标题，第 2 行列名
    .RegisterSheetReader(FixedTitle, new FixedTitleReader())
    .Build();

engine.ExportToFile(rows, "out.xlsx",
    new ExcelExportOptions { HeaderKind = FixedTitle, Columns = cols });
var back = engine.ImportFile<MyRow>("out.xlsx",
    new ExcelImportOptions { HeaderKind = FixedTitle, DataStartRowIndex = 2, Columns = cols });
```

> 注册到**新的 HeaderKind** 即为「追加」，注册到已存在的 `"Single"` / `"Group"` 即为「覆盖」默认实现。

### 6.6 关于读取快路径（A5/A6）

- 内置读取器在**访问器为 `ReflectionEntityAccessor` 且转换器为默认**时，会按列预解析「转换+赋值」编译委托并直接调用（省去每格字典查找），这是纯内部的性能优化。
- 当你替换 `IEntityAccessor` 或 `ICellValueConverter` 时，读取器**自动回退到接口路径**（`accessor.Write`），行为不受影响，只是少了快路径加速。
- 因此自定义访问器/转换器的扩展语义不变，无需感知快路径的存在。

---

## 7. 扩展的注意事项

### 7.1 写入器/读取器的表头布局必须一致

- 导出用 `HeaderKind = "X"` 写了某种布局，**导入必须用同一个 HeaderKind**，否则按错位读取。
- `ISheetWriter.HeaderRowCount` 必须与读取器假设的表头行数一致；导入时 `DataStartRowIndex` 也要与之对应（默认值只适配内置的单/分组表头）。
- 建议在自定义 `ISheetReader` 内按 `DataStartRowIndex` 计算数据起始行，不要硬编码行号。

### 7.2 样式不能跨工作簿复用

- `WriterStyles` 按 Workbook 在每次 `Write` 时创建，`ISheetWriter` 实现**不要**缓存任何 `ICellStyle`（NPOI 样式绑定所属 Workbook，跨簿复用会抛异常）。
- `WriterStyles` 内部会按「锁定/日期/交替行/边框/数字格式」组合缓存数据样式（ConcurrentDictionary，仅限当前 Workbook 内复用），避免为每列/每行重复建样式；表头按颜色缓存。
- 同理，自己写的写入策略若创建样式，请在方法内部基于传入的 `workbook` 创建。

### 7.3 扩展实现须保持无状态（可单例）

- `ExcelEngineBuilder` 默认把策略当单例用。自定义实现内**不要持有跨调用可变状态**；确有状态应放到每行/每次写入的局部对象中。
- 这是本库能「单例复用」的前提，也是并发/性能的基础。

### 7.4 整列锁定的实现细节

- 锁定铺底发生在**写单元格之前**（`ApplyColumnDefaults` → `SetDefaultColumnStyle`），因为 `SetDefaultColumnStyle` 会覆盖已存在单元格的样式。
- 数据单元格样式由 `WriterStyles.CellStyleFor` 按 `IsLocked` 决定，与默认列样式保持一致。
- 只有存在 `IsLocked` 列**或**显式传 `Protection` 时才执行 `ProtectSheet`；保护对**整张表**生效（锁与不锁靠单元格样式的 `IsLocked` 区分）。
- **排序/筛选与锁定**：`CT_SheetProtection` 的 `sort` / `autoFilter` 默认值为 `true`（表示禁止），若不显式放开，受保护工作表在 Excel 中 排序 和 自动筛选 会被禁用（与单元格是否锁定无关）。库默认放开（`AllowSort` / `AllowAutoFilter`），需要禁止时再置为 `false`。

### 7.5 公式占位符与列地址

- 默认解析器把 `{Field}` 替换为「当前行 + 该列」的单元格地址（如 `{Qty}` → `B3`）。占位符键是列的 `Field`，不是列名。
- 公式**无需等号**开头；`DefaultFormulaResolver` 只做花括号替换，不做公式合法性校验。
- 导出时引擎调用 `XSSFFormulaEvaluator.EvaluateAllFormulaCells` 预计算，因此导出的文件直接可见结果；导入时对公式单元格也会重新求值。

### 7.6 公式性能与求值策略

- **导出**：`EvaluateAllFormulaCells` 会对整簿所有公式单元格求值，NPOI 公式引擎较慢。大数据量 + 多公式列时可设 `ExcelExportOptions.CalculateFormulasOnExport = false` 跳过预计算，改为写 `calcPr.fullCalcOnLoad="1"`（打开时全量重算），由 Excel 打开时出结果，导出显著提速。代价：未用 Excel 打开前，其它程序读到的公式缓存值为空。
- **导入**：`DefaultCellValueConverter` 按**工作簿复用同一个公式求值器**（`ConditionalWeakTable<IWorkbook, IFormulaEvaluator>` 缓存），避免每个公式单元格新建求值器造成 O(n²) 劣化；求值器随工作簿 GC 回收，无泄漏。
- **注意**：NPOI 2.7.2 的 `XSSFWorkbook.SetForceFormulaRecalculation(bool)` 不能可靠写出 `fullCalcOnLoad`，库内直接操作 `CT_CalcPr.fullCalcOnLoad`（与 `SheetProtector` 写法一致）。

### 7.7 自定义值转换器的组合方式

- 推荐在自定义 `ICellValueConverter` 内部**组合** `DefaultCellValueConverter.Instance`（先取原始值再加工），避免重复实现类型转换逻辑。
- `IsNumeric` 返回 `true` 时，写入逻辑会走 `Convert.ToDouble` 数值分支；返回 `false` 时统一走 `ToCellText` 文本分支。需与列 `WriteAsNumeric` 配合理解。

### 7.8 读取器的属性类型解析

- 内置读取器通过 `entityType.GetProperty(field)` 推断目标属性类型，**找不到属性时按 `string` 处理**。
- 属性类型**在每个读取策略内按列预计算一次**（缓存于 `ConcurrentDictionary`，由 `ReflectionEntityAccessor.GetPropertyType` 提供），避免在「行 × 列」循环内重复反射。
- 因此用「字典访问器 + 字典行」这类无强类型属性场景，单元格数值会以字符串形态读回；如需类型化，请配套自定义 `ICellValueConverter` 或在访问器内自行转换。

### 7.9 编译式访问器与内存回收

- `ReflectionEntityAccessor` 的读写已编译为**表达式树委托**并按 `(Type, Field)` 静态缓存：每个组合只编译一次，之后复用。
- **无泄漏风险**：编译产物是托管委托（`DynamicMethod`），不产生动态程序集，随 GC 正常回收；静态缓存条数 = 使用过的 (Type, Field) 组合数，不随调用次数增长。
- 扩展作者若自行用表达式编译读写委托，请同样按 `(Type, Field)` 缓存，避免每次调用重复编译；**不要**用 `AssemblyBuilder` 动态程序集（.NET 中不可卸载，易造成内存泄漏）。

### 7.10 读/写热路径优化（不影响速度的改进）

- **A1 公式预检**：写数据前先扫描是否含公式列；无公式时不构建 `addressMap`，省去每行一次字典分配。
- **A2 列字母预计算**：导出前一次性算好列字母，不再每行重拼。
- **A3/D1 列定义缓存**：`AttributeColumnProvider` 按类型缓存，返回 `Clone()` 防污染（不影响速度，反而省去重复反射）。
- **A4 行对象工厂**：`ReflectionEntityAccessor.GetFactory` 缓存编译的 `new`，免每行 `Activator.CreateInstance`。
- **A5 读取快路径**：默认访问器时读取器按列预解析赋值委托（数组化，索引=列号），循环内直接调用，省去每格字典查找与二次转换。
  - 默认转换器 → **纯赋值委托**（值已为目标类型，省去每格一次 `ValueConvert`）；
  - 自定义转换器 → **转换+赋值委托**（仍省字典查找）；
  - 自定义访问器 → 自动回退接口路径。
- **A6 导出 getter 预解析**：默认访问器时导出按列预解析 getter 委托，省去每格 `GetOrAdd`（行类型从首元素解析，异构集合自动重解析）。
- **A7 列宽预计算**：表头字节长度、当前列宽按列预计算；内容宽度改用**原始值文本**（`value.ToString()`），避免对每个数值格调用 DataFormatter/`cell.ToString()`；公式列以表头宽度为准；日期用显示格式 `yyyy-MM-dd`。
- 以上改动均不改变默认行为与结果，压测确认：5000×30 导出 ≈1.04s（≈4800 行/s）/ 导入 ≈0.91s（≈5500 行/s），较上一版再提升约 20%；20000×30 导出 ≈4.3s / 导入 ≈4.1s。

### 7.11 泛型约束 `where T : new()`

- 引擎的导出/导入方法均要求 `T : new()`（读取时用 `Activator.CreateInstance` 建实体）。
- 匿名类型不可满足该约束；请使用具名类或程序化列 + 自定义访问器。

### 7.12 注册表键不区分大小写

- `SheetWriterRegistry` / `SheetReaderRegistry` 使用 `OrdinalIgnoreCase`，HeaderKind 大小写不敏感；`"group"` 与 `"Group"` 等价。

### 7.13 导入校验的语义

- 校验发生在读取策略内（`Readers/ColumnValidator.cs`），**单表头与分组表头都生效**，但各自按自身的列 key 约定（单表头用列名；分组表头用「列名_分组名」）。
- 「多余列」判定使用**全部**已知列（含隐藏列）作为识别集合，因此模板中的隐藏列不会被误判为多余列。
- 「只读列」仅指 `get-only` 属性（`CanWrite == false`）；`private set` 属性经反射可写，不在此列。
- 列 `Field` 在实体上不存在时按「可写」处理：这是为了兼容自定义访问器（如字典行）可处理任意字段；配合默认反射访问器时写入会被静默跳过。
- 校验结果**始终**写入 `ExcelImportOptions.Errors`（含已忽略列、重复列等提示），异常只按选项抛出，便于先收集后决策。

---

## 8. 扩展的边界

明确「本库不做什么 / 不能做什么」，避免扩展时踩坑：

### 8.1 文件格式边界

- **仅支持 `.xlsx`（XSSF/OpenXML）**，不支持 `.xls`（HSSF）。引擎内部固定 `XSSFWorkbook`。
- 输出为 ZIP 容器（魔数 `PK`），导入亦按 XSSF 解析。

### 8.2 表头布局边界

- 内置仅 `Single`（1 行）与 `Group`（2 行）两种布局。更多布局（三级表头、表头内嵌样式、行高表头等）**必须**自行实现 `ISheetWriter`/`ISheetReader` 并注册新 HeaderKind——这是本库的「扩展通道」，而非内置能力。

### 8.3 数据验证边界

- 支持：显式列表（`DropdownOptions`）、整数/小数区间、日期区间、自定义公式、公式列表（`ValidationKind` + `ValidationFormula1/2`）。
- 不支持：文本长度、时间（TIME）等更多类型；如需可在自定义写入策略中自行叠加 `IDataValidation`。
- 验证作用范围为 `ValidationMaxRow`（默认 1000），更大范围需调整该选项。

### 8.4 锁定与保护边界

- 保护粒度是**整个工作表**（`ProtectSheet`）；「列级锁定」靠单元格/默认列样式的 `IsLocked` 区分，不存在独立的按列保护对象。
- 密码为空（且未开 `ProtectWithoutPassword`）时 `SheetProtector.Apply` 直接返回（不保护）；开 `ProtectWithoutPassword` 后空密码会「保护但不设密码」。
- `AllowInsertRows=false` 对应 `CT_SheetProtection.insertRows = true`（即保持禁止）；`true` 则显式放开。该语义与 XML 属性命名相反，扩展时勿混淆。
- `sort` / `autoFilter` 与 `insertRows` 同语义：选项 `false` → 属性保持默认 `true`（禁止）；选项 `true` → 显式写 `0`（放开）。

### 8.5 公式边界

- 公式为**占位符替换**，不解析公式内容、不校验引用合法性；跨 Sheet 引用、易失函数等由 Excel/NPOI 自行处理。
- 默认解析器仅支持 `{Field}` 语法；其它语法需自定义 `IFormulaResolver`。
- 导入时对公式单元格求值依赖 NPOI 公式求值器；求值失败的公式回退为公式字符串本身。

### 8.6 空行/类型边界

- `SkipEmptyRows` 的判定是「任意单元格有非空值即视为有数据」。对**含数值类型属性的实体**，全空行的数值列默认为 `0`，会被视为有意义而导入（这是当前实现的行为边界，非缺陷）。
- 若需严格跳过，可自行在读取策略或访问器中按业务规则过滤。

### 8.7 导入校验边界

- 校验只针对**列名对应关系**，不校验单元格值与属性类型的匹配（如数值列写入文本不会报错，仅按转换规则尽力转换）。
- 校验基于「表头文字」而非列位置，不强制列顺序与 `Order` 一致。
- 「只读」以反射的 `CanWrite` 为准，不感知 `init` 等编译期只读语义。
- 自定义访问器（字典行等）不经过属性反射校验，其列的读写完全由访问器决定。

### 8.8 单 Sheet 导入边界

- `Import<T>` 一次解析一个 Sheet（按 `SheetName`/`SheetIndex`）。多 Sheet 解析需多次调用 `Import<T>`。
- `ImportAll<T>` 一次解析全部 Sheet，但**要求所有 Sheet 结构一致（同一 `T`）**；结构不同的 Sheet 会按缺列报错。
- 多 Sheet 导出支持异类型（`SheetSpec`），但同一 Sheet 内仍是同一 `RowType`。

### 8.9 并发与线程安全边界

- 引擎与内置策略**无共享可变状态**，可跨线程复用引擎实例。
- 自定义扩展若引入可变状态，线程安全由扩展作者负责（见 7.3）。

### 8.10 对库本身修改的边界（开闭原则的代价）

- 引擎的流水线（列解析 → 策略选择 → 生成/解析）固定，**不对外暴露 Sheet 级中间态**；要插入额外处理（如合并单元格、图片、批注、超链接）需在自定义 `ISheetWriter` 内完成，或扩展引擎（不在本库既定职责内）。
- 想为库新增「内置能力」时，优先考虑以**扩展实现**提交，而非修改核心类，以保持对修改关闭。

---

## 9. 构建与测试

```bash
# 还原 + 构建整个解决方案
dotnet build MagicFrame.Excel.sln

# 运行测试（70 个用例）
dotnet test MagicFrame.Excel.sln

# 运行演示（输出到 bin/.../output/，可传第二个参数指定输出目录）
dotnet run --project demos/MagicFrame.Excel.Demo                    # 全部
dotnet run --project demos/MagicFrame.Excel.Demo -- basic           # 仅基础
dotnet run --project demos/MagicFrame.Excel.Demo -- ext             # 仅扩展
dotnet run --project demos/MagicFrame.Excel.Demo -- basic D:\out    # 指定输出目录

# 压力测试（独立于单元测试的可执行项目，tools/）
dotnet run --project tools/MagicFrame.Excel.StressTest             # 默认 5000 行 × 30 列 × 3 次
dotnet run --project tools/MagicFrame.Excel.StressTest -- 10000 30 3   # 行数 列数 次数（列数≤30）
dotnet run --project tools/MagicFrame.Excel.StressTest -- formula 5000 30  # 公式导出：预计算 vs 打开时重算

# 打包 NuGet（Release，产出 nupkg + snupkg）
dotnet pack MagicFrame.Excel.csproj -c Release -o ./artifacts
```

- **CI**：`.github/workflows/ci.yml`（GitHub Actions）——push/PR 自动 build + test + 压测冒烟；打 `v*` tag 时额外 `dotnet pack` 并上传产物。

### 9.1 压力测试说明

- 独立可执行项目 `tools/MagicFrame.Excel.StressTest`，与单元测试完全隔离。
- 实体 `StressRow` 含 30 个属性（10 字符串 / 10 数值 / 5 金额 / 3 日期 / 2 布尔），覆盖常用单元格类型路径。
- 流程：先**预热**一次（触发 JIT 与表达式编译缓存，不计入统计），再迭代 N 次；每次校验行数与抽样值，最后输出导出/导入平均耗时与吞吐（行/s）。
- 参考数据（本机 net6.0 / Debug，经热路径优化后）：

| 规模 | 导出 | 导入 | 文件 |
| --- | --- | --- | --- |
| 5000 × 30 | ≈ 1.04 s（≈4800 行/s） | ≈ 0.91 s（≈5500 行/s） | ≈ 0.84 MB |
| 20000 × 30 | ≈ 4.3 s（≈4700 行/s） | ≈ 4.1 s（≈4900 行/s） | ≈ 3.4 MB |

- **公式导出对比**（`-- formula 5000 30`，31 列含 1 个公式列）：

| 模式 | 导出 | 导入 | 文件 |
| --- | --- | --- | --- |
| 预计算（默认） | ≈ 1.22 s | ≈ 1.57 s | ≈ 0.92 MB |
| 打开时重算（`CalculateFormulasOnExport=false`） | ≈ 1.15 s | ≈ 1.23 s | ≈ 0.89 MB |

> 说明：公式导入耗时包含对每个公式单元格的求值。库内已按工作簿复用求值器，使导入随行数**线性**增长（修复前每个公式单元格新建求值器会退化为 O(n²)，5000 行公式导入耗时超过 5 分钟）。

### 9.2 内存与 CPU 说明

**内存（不是泄漏）**：压测（8000×30 = 24 万单元格）实测：迭代期间托管内存 ≈130~230 MB、进程 WorkingSet ≈290~400 MB；**`GC.Collect()` 后回落至 ≈7~11 MB / ≈100~230 MB**，说明全部可回收，无泄漏。

- 峰值来自 **NPOI XSSF 的整簿驻留模型**（每个单元格的 XSSFCell 包装 + CT_Xml 对象树），8000×30 即 24 万个单元格对象，属 XSSF（非流式）的固有开销。
- 便捷方法（`ExportBytes`/`ExportBase64`/`ExportToFile`）**刻意不主动 Dispose 工作簿**：实测主动 Dispose 会使 5000×30 导出从 ≈1.05s 升到 ≈1.5s（GC 在导出窗口内回收上一簿导致停顿），而内存本就由 GC 完全回收，故选择「不拖慢导出」。
- WorkingSet 高企是 Windows 不会主动收缩已触碰的物理内存页，属正常现象，不代表占用未释放。
- 需要极致省内存时：用 `Export()` 拿到结果后自行 `Dispose()`，或分批导出/接入流式（见 8.x 暂缓项 B1 SXSSF）。
- `ExcelExportResult.Base64` 为**惰性计算**，仅取 `Bytes` 时不额外生成 Base64 字符串。

**CPU（≈1 核，属正常）**：NPOI 为**单线程**库，18% 左右 ≈ 满负荷占用 1 个核心（多核机器上单线程基准的正常表现）。导出 CPU 主要花在：创建 24 万单元格对象 + 序列化 XML + ZIP 压缩；导入花在：解析 + 单元格对象化。库自身已通过编译式委托/按列预解析把业务层开销降到最低，剩余为 NPOI 固有成本；进一步提速需流式（SXSSF 导出 / SAX 导入，暂缓项）。

---

## 10. 相关图表（drawio）

`docs/diagrams/` 下提供可编辑的 drawio 图（用 https://app.diagrams.net 或 VS Code 插件打开）：

| 文件 | 内容 |
| --- | --- |
| `architecture.drawio` | 系统总体架构（分层：应用 / 引擎 / 扩展接口 / 默认实现 / NPOI） |
| `export-flow.drawio` | 导出流程（含表头策略分支、公式处理策略分支） |
| `import-flow.drawio` | 导入流程（含列对应校验、空行跳过、逐行解析） |
| `extension-points.drawio` | 扩展点与开闭原则（6 个扩展点：默认 vs 自定义） |
| `data-model.drawio` | 核心数据模型（ExcelColumn 单一事实来源 + 选项类） |
| `performance.drawio` | 关键性能与保护机制（编译式访问器 / 求值器复用 / 公式延后 / 锁定放开排序筛选） |

---

## 附：API 速查

```
ExcelEngine.CreateDefault() / ExcelEngineBuilder.Create()...Build()
  ├─ Export<T> / ExportBytes<T> / ExportBase64<T> / ExportStream<T> / ExportToFile<T>
  ├─ ExportSheets(SheetSpec[]) / ExportSheets(IDictionary<string, IList<T>>)
  └─ Import<T>(bytes | base64 | file | stream) / ImportAll<T>(bytes)

ExcelExportOptions   { SheetName, HeaderKind, Columns, Protection, FreezeHeader, FreezeColumns,
                       AutoFilter, RowHeight, AlternateRowFillColor, EnableBorders, Progress,
                       ValidationMaxRow, MaxColumnWidth, CalculateFormulasOnExport }
ExcelImportOptions   { SheetName, SheetIndex, HeaderKind, DataStartRowIndex, SkipEmptyRows,
                       ThrowOnMissingColumns, ThrowOnUnexpectedColumns, IgnoreReadOnlyColumns,
                       CollectRowErrors, SourceRowProperty, Progress, Columns, Errors, Issues }
SheetProtectionOptions { Enabled, Password, ProtectWithoutPassword,
                         AllowInsertRows, AllowDeleteRows, AllowFormat, AllowSort, AllowAutoFilter }
ExcelColumn / [ExcelColumn]  { Name, Field, Order, Visible, IsLocked, Formula, ZeroShowWhiteSpace,
                               HeaderColor, GroupName, GroupText, DropdownOptions, ValidationKind,
                               ValidationFormula1/2, Width, Alignment, WriteAsNumeric, NumberFormat,
                               ValueMap（含 unknown 哨兵键）, ValueMappings }
ValidationKind   { None, List, Integer, Decimal, Date, CustomFormula, FormulaList }
ImportIssue      { Severity(Info|Warning|Error), Code, Message, RowNumber }
HeaderKinds          { Single, Group }
```

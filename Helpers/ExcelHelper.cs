using System.Globalization;
using System.IO;

namespace ExcelMaker.Helpers;

/// <summary>
/// Excel 读写辅助：功能2（库位名称导入）与功能3（库存明细处理）。
/// 使用 ClosedXML（MIT），完全复制单元格“内容”，格式按模板 B 的列定义单独设置。
/// </summary>
public static class ExcelHelper
{
    #region 功能2：库位名称表头探测 + 读取

    private static readonly Dictionary<string, Action<StockNameRow, string>> StockNameMap = new()
    {
        ["公司代码"] = (r, v) => r.CompanyCode = v,
        ["工厂代码"] = (r, v) => r.FactoryCode = v,
        ["库位代码"] = (r, v) => r.StockCode = v,
        ["索引标识"] = (r, v) => r.StockIndex = v,
        ["库位名称"] = (r, v) => r.StockName = v,
    };

    /// <summary>
    /// 读取第一个工作表，逐行向下查找表头（可能不在第一行），
    /// 找到后从下一行开始读取到连续空行。
    /// </summary>
    public static List<StockNameRow> ReadStockNameSheet(string path)
    {
        // 先判空：path 为 null/空时 new XLWorkbook(null) 会抛 ArgumentNullException（表现为“报错 null”）
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("未提供 Excel 文件路径（path 为空），无法读取。请确认调用时传入了有效的文件完整路径。", nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"未找到 Excel 文件：{path}", path);

        // ClosedXML 仅支持 .xlsx（Open XML）；旧版 .xls 无法读取，会表现为工作簿为空/抛异常
        if (Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("不支持旧版 .xls 格式，请用 Excel 另存为 .xlsx 后再导入（ClosedXML 仅支持 .xlsx）");

        XLWorkbook wb;
        try
        {
            wb = new XLWorkbook(path);
        }
        catch (Exception ex)
        {
            // 把格式不符/文件损坏等底层异常转成清晰中文提示（包含原始信息便于排查）
            throw new InvalidOperationException($"打开 Excel 失败，请确认文件是 .xlsx 格式且未损坏：{ex.Message}", ex);
        }

        using (wb)
        {
            if (wb.Worksheets.Count == 0)
                throw new InvalidOperationException("工作簿中没有任何工作表（请确认文件是 .xlsx 格式，且未被损坏）");

            var ws = wb.Worksheets.Worksheet(1);
            var lastRowRange = ws.LastRowUsed();
            var lastRow = lastRowRange == null ? 0 : lastRowRange.RowNumber();
            if (lastRow == 0)
                throw new InvalidOperationException("第一个工作表没有任何数据行（请确认数据在第一个工作表，且表头/数据未被清空）");

        // 1) 找到表头行
        int headerRow = -1;
        Dictionary<string, int>? columnMap = null;
        var lastCellUsed = ws.LastCellUsed();
        int lastCol = lastCellUsed == null ? 50 : lastCellUsed.Address.ColumnNumber;
        for (int r = 1; r <= Math.Min(lastRow, 100); r++)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= lastCol; c++)
            {
                var text = ws.Cell(r, c).GetString().Trim();
                if (StockNameMap.ContainsKey(text))
                    map[text] = c;
            }
            if (map.Count == StockNameMap.Count) { headerRow = r; columnMap = map; break; }
        }
        if (headerRow < 0 || columnMap == null)
            throw new InvalidOperationException("未找到表头（公司代码/工厂代码/库位代码/索引标识/库位名称）");

        // 2) 读取数据行
        var rows = new List<StockNameRow>();
        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var row = new StockNameRow();
            bool any = false;
            foreach (var kv in columnMap)
            {
                var v = ws.Cell(r, kv.Value).GetString().Trim();
                StockNameMap[kv.Key](row, v);
                if (v.Length > 0) any = true;
            }
            if (!any) break; // 遇到整行空白即停止
            rows.Add(row);
        }
        return rows;
        }
    }

    #endregion

    #region 功能3：库存明细（A 表读取 + B 表生成）

    // A 表需要读取的列（字母）
    private static readonly string[] InventoryColumns = { "B", "D", "E", "F", "G", "H", "K", "L", "M", "N", "R" };

    /// <summary>
    /// 读取 A 表第一个工作表（表头第1行，内容第2行起），返回每行按列字母索引的值。
    /// </summary>
    public static List<Dictionary<string, string>> ReadInventorySheet(string path)
    {
        // 先判空：path 为 null/空时 new XLWorkbook(null) 会抛 ArgumentNullException（表现为“报错 null”）
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("未提供 Excel 文件路径（path 为空），无法读取。请确认调用时传入了有效的文件完整路径。", nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"未找到 Excel 文件：{path}", path);

        // ClosedXML 仅支持 .xlsx；旧版 .xls 无法读取，会表现为工作簿为空/抛异常
        if (Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("不支持旧版 .xls 格式，请用 Excel 另存为 .xlsx 后再导入（ClosedXML 仅支持 .xlsx）");

        XLWorkbook wb;
        try
        {
            wb = new XLWorkbook(path);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"打开 Excel 失败，请确认文件是 .xlsx 格式且未损坏：{ex.Message}", ex);
        }

        using (wb)
        {
            if (wb.Worksheets.Count == 0)
                throw new InvalidOperationException("工作簿中没有任何工作表（请确认文件是 .xlsx 格式，且未被损坏）");

            var ws = wb.Worksheets.Worksheet(1);
            var lastRowRange2 = ws.LastRowUsed();
            var lastRow = lastRowRange2 == null ? 0 : lastRowRange2.RowNumber();
            if (lastRow < 2)
                throw new InvalidOperationException("第一个工作表没有足够的数据行（A 表表头应在第1行，数据从第2行起）");

        var rows = new List<Dictionary<string, string>>();
        for (int r = 2; r <= lastRow; r++)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool any = false;
            foreach (var c in InventoryColumns)
            {
                var v = ws.Cell($"{c}{r}").GetString().Trim();
                dict[c] = v;
                if (v.Length > 0) any = true;
            }
            if (!any) continue; // 跳过整行空白
            rows.Add(dict);
        }
        return rows;
        }
    }

    /// <summary>
    /// 基于空白模板 B 生成库存明细成品并保存为新文件（模板不会被覆盖）。
    /// </summary>
    public static void BuildInventoryWorkbook(
        string templatePath,
        string outputPath,
        List<Dictionary<string, string>> aRows,
        Dictionary<string, (string StockIndex, string StockName)> stockDict,
        Dictionary<string, string> authorDict)
    {
        if (aRows.Count == 0)
            throw new InvalidOperationException("未读取到库存明细数据行（A 表第2行起无内容）");

        using var wb = new XLWorkbook(templatePath);
        var ws = wb.Worksheets.Worksheet(1);

        // 列样式定义（微软雅黑，按需求逐项设置，未指定的属性保持默认）
        var sA = Style(9, XLAlignmentVerticalValues.Center, null, null, false);
        var sB = Style(9, XLAlignmentVerticalValues.Bottom, XLAlignmentHorizontalValues.Center, null, true);
        var sC = Style(9, XLAlignmentVerticalValues.Center, XLAlignmentHorizontalValues.Left, null, true);
        var sD = sC; var sE = sC; var sF = sC; var sG = sC;
        var sH = Style(9, XLAlignmentVerticalValues.Bottom, XLAlignmentHorizontalValues.Left, null, true);
        var sI = Style(9, XLAlignmentVerticalValues.Bottom, null, null, true);
        var sJ = Style(9, XLAlignmentVerticalValues.Bottom, XLAlignmentHorizontalValues.Center, "0.00%", true);
        var sK = Style(11, XLAlignmentVerticalValues.Center, XLAlignmentHorizontalValues.Center, null, true);
        var sL = Style(9, XLAlignmentVerticalValues.Center, XLAlignmentHorizontalValues.Left, "#,##0.00", true);
        var sM = sC;
        var sN = Style(9, XLAlignmentVerticalValues.Center, XLAlignmentHorizontalValues.Center, "_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * \"-\"??_ ;_ @_", true);
        var sO = Style(9, XLAlignmentVerticalValues.Center, XLAlignmentHorizontalValues.Left, "_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * \"-\"??_ ;_ @_", true);
        var sP = sO; var sQ = sO; var sR = sO; var sS = sO; var sT = sO;
        var sU = Style(9, XLAlignmentVerticalValues.Bottom, null, "#,##0.00", true);
        var sV = sU;
        var sW = Style(9, XLAlignmentVerticalValues.Bottom, XLAlignmentHorizontalValues.Left, "#,##0.00", true);
        var sX = sW; var sY = sW; var sZ = sW; var sAA = sW; var sAB = sW;
        var sAC = sU; var sAD = sU;
        var sAE = Style(9, XLAlignmentVerticalValues.Bottom, null, "_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * \"-\"??_ ;_ @_", true);

        for (int k = 0; k < aRows.Count; k++)
        {
            int bRow = 9 + k;
            var a = aRows[k];
            string Get(string col) => a.TryGetValue(col, out var v) ? v : "";

            string splant = Get("D");
            string stock = Get("G");
            string byd = Get("E");

            // 字典查表
            stockDict.TryGetValue($"{splant}|{stock}", out var sn);
            authorDict.TryGetValue($"{splant}|{stock}|{byd}", out var author);

            // A: 公式 =D&G&E
            SetCell(ws, bRow, "A", null, $"D{bRow}&G{bRow}&E{bRow}", sA);
            // B: 序号 = A表B列
            SetCell(ws, bRow, "B", Get("B"), null, sB);
            // C: = A表F列
            SetCell(ws, bRow, "C", Get("F"), null, sC);
            // D: = A表H列
            SetCell(ws, bRow, "D", Get("H"), null, sD);
            // E: = A表L列
            SetCell(ws, bRow, "E", Get("L"), null, sE);
            // F: = A表M列
            SetCell(ws, bRow, "F", Get("M"), null, sF);
            // G: = A表K列
            SetCell(ws, bRow, "G", Get("K"), null, sG);
            // H: FACTORY_STOCK_NAME.STOCK_INDEX
            SetCell(ws, bRow, "H", sn.StockIndex, null, sH);
            // I: FACTORY_STOCK_NAME.STOCK_NAME
            SetCell(ws, bRow, "I", sn.StockName, null, sI);
            // J: 固定文本“是”（百分比格式）
            SetCell(ws, bRow, "J", "是", null, sJ);
            // K: 复盘人（D_S_INVENTORY_DETAIL 最新一笔 author）
            SetCell(ws, bRow, "K", author, null, sK);
            // L: 空（数值）
            SetCell(ws, bRow, "L", null, null, sL);
            // M: = A表N列
            SetCell(ws, bRow, "M", Get("N"), null, sM);
            // N: 公式 =P/O
            SetCell(ws, bRow, "N", null, $"P{bRow}/O{bRow}", sN);
            // O: = A表R列
            SetCell(ws, bRow, "O", Get("R"), null, sO);
            // P: 空（会计专用）
            SetCell(ws, bRow, "P", null, null, sP);
            // Q: = A表R列
            SetCell(ws, bRow, "Q", Get("R"), null, sQ);
            // R: 空
            SetCell(ws, bRow, "R", null, null, sR);
            // S: = A表R列
            SetCell(ws, bRow, "S", Get("R"), null, sS);
            // T: 空
            SetCell(ws, bRow, "T", null, null, sT);
            // U: 空（数值）
            SetCell(ws, bRow, "U", null, null, sU);
            // V: 空（数值）
            SetCell(ws, bRow, "V", null, null, sV);
            // W: 公式 =IF(S="","",IF(S>O,S-O,0))
            SetCell(ws, bRow, "W", null, $"IF(S{bRow}=\"\",\"\",IF(S{bRow}>O{bRow},S{bRow}-O{bRow},0))", sW);
            // X: 公式 =IF(S="","",N*W)
            SetCell(ws, bRow, "X", null, $"IF(S{bRow}=\"\",\"\",N{bRow}*W{bRow})", sX);
            // Y: 公式 =IF(S="","",IF(S<O,O-S,0))
            SetCell(ws, bRow, "Y", null, $"IF(S{bRow}=\"\",\"\",IF(S{bRow}<O{bRow},O{bRow}-S{bRow},0))", sY);
            // Z: 公式 =IF(S="","",Y*N)
            SetCell(ws, bRow, "Z", null, $"IF(S{bRow}=\"\",\"\",Y{bRow}*N{bRow})", sZ);
            // AA: 公式 =IF(S="","",X-Z)
            SetCell(ws, bRow, "AA", null, $"IF(S{bRow}=\"\",\"\",X{bRow}-Z{bRow})", sAA);
            // AB: 公式 =IF(S="","",X+Z)
            SetCell(ws, bRow, "AB", null, $"IF(S{bRow}=\"\",\"\",X{bRow}+Z{bRow})", sAB);
            // AC: 空（数值）
            SetCell(ws, bRow, "AC", null, null, sAC);
            // AD: 空（数值）
            SetCell(ws, bRow, "AD", null, null, sAD);
            // AE: 空（会计专用）
            SetCell(ws, bRow, "AE", null, null, sAE);
        }

        // 汇总行：P 列 = SUM(P9:P{末数据行})
        int lastDataRow = 8 + aRows.Count;   // 第9行起，共 aRows.Count 行
        int summaryRow = lastDataRow + 1;
        SetCell(ws, summaryRow, "P", null, $"SUM(P9:P{lastDataRow})", sP);

        // 保存前重算所有公式并写入缓存值，避免 WPS/Excel 打开时公式不计算（需双击单元格才显示）
        wb.RecalculateAllFormulas();
        wb.SaveAs(outputPath);
    }

    #endregion

    #region 内部工具

    private static ColumnStyle Style(int fontSize, XLAlignmentVerticalValues? v, XLAlignmentHorizontalValues? h, string? fmt, bool border)
        => new() { FontSize = fontSize, Vertical = v, Horizontal = h, NumberFormat = fmt, BorderAll = border };

    private sealed class ColumnStyle
    {
        public int FontSize = 9;
        public string FontName = "微软雅黑";
        public XLAlignmentVerticalValues? Vertical;
        public XLAlignmentHorizontalValues? Horizontal;
        public string? NumberFormat; // null = 常规
        public bool BorderAll;
    }

    private static void SetCell(IXLWorksheet ws, int row, string col, string? value, string? formula, ColumnStyle s)
    {
        var cell = ws.Cell($"{col}{row}");
        cell.Style.Font.FontName = s.FontName;
        cell.Style.Font.FontSize = s.FontSize;
        if (s.Vertical.HasValue) cell.Style.Alignment.Vertical = s.Vertical.Value;
        if (s.Horizontal.HasValue) cell.Style.Alignment.Horizontal = s.Horizontal.Value;
        if (s.NumberFormat != null) cell.Style.NumberFormat.Format = s.NumberFormat;
        if (s.BorderAll) cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        if (formula != null)
        {
            cell.FormulaA1 = formula;
        }
        else if (!string.IsNullOrEmpty(value))
        {
            // 数值/会计专用列：若内容可解析为数字则写入数值（保证公式可计算）；
            // 常规列（代码类）：原样写入字符串，保留前导撇号与前导零。
            bool numericStyle = s.NumberFormat != null &&
                                (s.NumberFormat.Contains('#') || s.NumberFormat.Contains('0'));
            if (numericStyle && double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                cell.Value = d;
            else
                cell.Value = value; // 含前导撇号 ' 的文本也会被原样保留
        }
    }

    #endregion
}

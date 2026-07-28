namespace ExcelMaker.Services;

/// <summary>
/// 数据库访问（Oracle），连接串来自 db.ini 中 AES 加密的密文，
/// 与 AutoSteelCheckMain 保持一致的解析方式。
/// </summary>
public class DatabaseService
{
    private readonly ConfigService _config;
    private readonly CryptoService _crypto;
    private string? _connectionString;

    public DatabaseService(ConfigService config, CryptoService crypto)
    {
        _config = config;
        _crypto = crypto;
    }

    private string GetConnectionString()
    {
        if (_connectionString == null)
        {
            var cipher = _config.ConnectionStringCipher;
            if (string.IsNullOrEmpty(cipher) || cipher == "REPLACE_WITH_ENCRYPTED_STRING")
                throw new InvalidOperationException("数据库连接串未配置，请先运行加密工具生成密文");

            var plain = _crypto.Decrypt(cipher);
            var poolParams = "Connection Lifetime=300;Min Pool Size=2;Max Pool Size=10;Incr Pool Size=2;Decr Pool Size=1;";
            _connectionString = plain.Contains("Pool Size", StringComparison.OrdinalIgnoreCase)
                ? plain : $"{plain};{poolParams}";
        }
        return _connectionString;
    }

    private async Task<OracleConnection> OpenConnectionAsync(CancellationToken ct = default)
    {
        var conn = new OracleConnection(GetConnectionString());
        await conn.OpenAsync(ct);
        return conn;
    }

    // ── 登录 ──
    public async Task<bool> ValidateLoginAsync(string username, string md5Password)
    {
        const string sql = "SELECT COUNT(1) FROM usr_info WHERE usrname = :u AND md5_pwd = :p";
        await using var conn = await OpenConnectionAsync();
        await using var cmd = new OracleCommand(sql, conn);
        cmd.Parameters.Add(new OracleParameter("u", username));
        cmd.Parameters.Add(new OracleParameter("p", md5Password));
        return Convert.ToInt32(await cmd.ExecuteScalarAsync()) > 0;
    }

    public async Task<string> GetUserFullNameAsync(string username)
    {
        const string sql = "SELECT allname FROM usr_info WHERE usrname = :u";
        await using var conn = await OpenConnectionAsync();
        await using var cmd = new OracleCommand(sql, conn);
        cmd.Parameters.Add(new OracleParameter("u", username));
        var result = await cmd.ExecuteScalarAsync();
        return result?.ToString() ?? username;
    }

    // ── 功能2：导入更新库位名称（MERGE）──
    /// <summary>
    /// 将读取到的库位名称行 MERGE 进 FACTORY_STOCK_NAME。
    /// 匹配条件：splant_no + stock_no；更新 AUTHOR（登录人工号）与 UPDATE_TIME（SYSDATE）。
    /// </summary>
    public async Task<int> MergeStockNameAsync(IReadOnlyList<StockNameRow> rows, string author)
    {
        if (rows.Count == 0) return 0;

        const string mergeSql = @"
MERGE INTO FACTORY_STOCK_NAME t
USING (SELECT :p_fc  AS factory_code,
              :p_sp  AS splant_no,
              :p_sn  AS stock_no,
              :p_si  AS stock_index,
              :p_snm AS stock_name
       FROM dual) s
ON (t.splant_no = s.splant_no AND t.stock_no = s.stock_no)
WHEN MATCHED THEN
  UPDATE SET t.stock_name  = s.stock_name,
             t.stock_index = s.stock_index,
             t.factory_code = s.factory_code,
             t.author       = :p_author,
             t.update_time  = SYSDATE
WHEN NOT MATCHED THEN
  INSERT (factory_code, splant_no, stock_no, stock_index, stock_name, author, update_time)
  VALUES (s.factory_code, s.splant_no, s.stock_no, s.stock_index, s.stock_name, :p_author, SYSDATE)";

        await using var conn = await OpenConnectionAsync();
        using var tx = conn.BeginTransaction();

        await using var cmd = new OracleCommand(mergeSql, conn, tx);
        cmd.BindByName = true;
        cmd.Parameters.Add(new OracleParameter("p_fc", OracleDbType.Varchar2));
        cmd.Parameters.Add(new OracleParameter("p_sp", OracleDbType.Varchar2));
        cmd.Parameters.Add(new OracleParameter("p_sn", OracleDbType.Varchar2));
        cmd.Parameters.Add(new OracleParameter("p_si", OracleDbType.Varchar2));
        cmd.Parameters.Add(new OracleParameter("p_snm", OracleDbType.Varchar2));
        cmd.Parameters.Add(new OracleParameter("p_author", OracleDbType.Varchar2));

        var processed = 0;
        foreach (var r in rows)
        {
            cmd.Parameters["p_fc"].Value = (object?)r.CompanyCode?.Trim() ?? DBNull.Value;
            cmd.Parameters["p_sp"].Value = (object?)r.FactoryCode?.Trim() ?? DBNull.Value;
            cmd.Parameters["p_sn"].Value = (object?)r.StockCode?.Trim() ?? DBNull.Value;
            cmd.Parameters["p_si"].Value = (object?)r.StockIndex?.Trim() ?? DBNull.Value;
            cmd.Parameters["p_snm"].Value = (object?)r.StockName?.Trim() ?? DBNull.Value;
            cmd.Parameters["p_author"].Value = (object?)author?.Trim() ?? DBNull.Value;
            await cmd.ExecuteNonQueryAsync();
            processed++;
        }

        tx.Commit();
        return processed;
    }

    // ── 功能3：库位索引/名称字典（按 splant_no + stock_no）──
    public async Task<Dictionary<string, (string StockIndex, string StockName)>> GetStockNameLookupAsync()
    {
        const string sql = @"
SELECT splant_no, stock_no, stock_index, stock_name
FROM FACTORY_STOCK_NAME";

        var dict = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        await using var conn = await OpenConnectionAsync();
        await using var cmd = new OracleCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var splant = reader.IsDBNull(0) ? "" : reader.GetString(0);
            var stock = reader.IsDBNull(1) ? "" : reader.GetString(1);
            var index = reader.IsDBNull(2) ? "" : reader.GetString(2);
            var name = reader.IsDBNull(3) ? "" : reader.GetString(3);
            dict[$"{splant}|{stock}"] = (index, name);
        }
        return dict;
    }

    // ── 功能3：复盘人字典（按 FACTORY_NO + STOCK_NO + SAP_BYDID，取选中年月最新一笔）──
    public async Task<Dictionary<string, string>> GetInventoryAuthorLookupAsync(string yyyymm)
    {
        const string sql = @"
SELECT FACTORY_NO, STOCK_NO, SAP_BYDID, author FROM (
  SELECT FACTORY_NO, STOCK_NO, SAP_BYDID, author,
         ROW_NUMBER() OVER (PARTITION BY FACTORY_NO, STOCK_NO, SAP_BYDID ORDER BY UPDATE_TIME DESC) rn
  FROM D_S_INVENTORY_DETAIL
  WHERE TO_CHAR(UPDATE_TIME, 'YYYYMM') = :ym
) WHERE rn = 1";

        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        await using var conn = await OpenConnectionAsync();
        await using var cmd = new OracleCommand(sql, conn);
        cmd.BindByName = true;
        cmd.Parameters.Add(new OracleParameter("ym", yyyymm.Trim()));
        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var factory = reader.IsDBNull(0) ? "" : reader.GetString(0);
            var stock = reader.IsDBNull(1) ? "" : reader.GetString(1);
            var byd = reader.IsDBNull(2) ? "" : reader.GetString(2);
            var author = reader.IsDBNull(3) ? "" : reader.GetString(3);
            dict[$"{factory}|{stock}|{byd}"] = author;
        }
        return dict;
    }

    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            await using var conn = await OpenConnectionAsync();
            await using var cmd = new OracleCommand("SELECT 1 FROM DUAL", conn);
            await cmd.ExecuteScalarAsync();
            return true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "数据库连接测试失败");
            return false;
        }
    }

    public void ClearPool()
    {
        try { OracleConnection.ClearAllPools(); }
        catch { }
    }
}

/// <summary>
/// 库位名称行（公司代码/工厂代码/库位代码/索引标识/库位名称）
/// </summary>
public class StockNameRow
{
    public string? CompanyCode { get; set; }
    public string? FactoryCode { get; set; }
    public string? StockCode { get; set; }
    public string? StockIndex { get; set; }
    public string? StockName { get; set; }
}

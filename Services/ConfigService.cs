namespace ExcelMaker.Services;

public class ConfigService
{
    private const string ConfigDir = @"D:\ExcelMaker";
    private readonly string _configIniPath;
    private readonly string _dbIniPath;

    public ConfigService()
    {
        _configIniPath = Path.Combine(ConfigDir, "config.ini");
        _dbIniPath = Path.Combine(ConfigDir, "db.ini");
    }

    public bool Exists() => File.Exists(_configIniPath) && File.Exists(_dbIniPath);

    public List<string> MissingFiles()
    {
        var missing = new List<string>();
        if (!File.Exists(_configIniPath)) missing.Add(_configIniPath);
        if (!File.Exists(_dbIniPath)) missing.Add(_dbIniPath);
        return missing;
    }

    // ── db.ini ──
    public string ConnectionStringCipher => ReadDb("database", "ConnectionStringCipher", "");

    // ── config.ini ──
    // 库存明细模板（空白模板，每次导入时读取，不会被覆盖）
    public string TemplatePath => ReadConfig("excel", "TemplatePath", @"D:\ExcelTemp\Template.xlsx");
    // 导出目录（处理库存明细后的成品默认保存位置）
    public string OutputDir => ReadConfig("excel", "OutputDir", @"D:\ExcelTemp");

    public void EnsureDirectories()
    {
        if (!Directory.Exists(ConfigDir))
            Directory.CreateDirectory(ConfigDir);
        var logDir = Path.Combine(ConfigDir, "logs");
        if (!Directory.Exists(logDir))
            Directory.CreateDirectory(logDir);
    }

    // ── 底层读取 ──
    private string ReadConfig(string section, string key, string defaultValue)
        => IniHelper.Read(_configIniPath, section, key, defaultValue);

    private string ReadDb(string section, string key, string defaultValue)
        => IniHelper.Read(_dbIniPath, section, key, defaultValue);
}

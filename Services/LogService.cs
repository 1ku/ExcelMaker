namespace ExcelMaker.Services;

public class LogService
{
    private const string LogDir = @"D:\ExcelMaker\logs";
    private bool _initialized;

    public void Initialize()
    {
        if (_initialized) return;

        if (!Directory.Exists(LogDir))
            Directory.CreateDirectory(LogDir);

        var logPath = Path.Combine(LogDir, "ExcelMaker_.log");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(logPath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                encoding: Encoding.UTF8)
            .CreateLogger();

        _initialized = true;
    }
}

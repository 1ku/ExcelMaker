using Microsoft.Extensions.DependencyInjection;

namespace ExcelMaker;

public partial class App : Application
{
    public static IServiceProvider ServiceProvider { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();

        services.AddSingleton<ConfigService>();
        services.AddSingleton<CryptoService>();
        services.AddSingleton<LogService>();
        services.AddSingleton<DatabaseService>();
        services.AddTransient<LoginViewModel>();
        services.AddSingleton<MainViewModel>();

        ServiceProvider = services.BuildServiceProvider();

        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            var ex = args.ExceptionObject as Exception;
            Log.Fatal(ex, "未处理的异常");
            MessageBox.Show($"程序发生严重错误，即将退出。\n{ex?.Message}", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            Environment.Exit(1);
        };

        var logService = ServiceProvider.GetRequiredService<LogService>();
        logService.Initialize();
        Log.Information("ExcelMaker 启动");

        var config = ServiceProvider.GetRequiredService<ConfigService>();
        if (!config.Exists())
        {
            Log.Error("配置文件缺失: {Files}", string.Join(", ", config.MissingFiles()));
            MessageBox.Show(
                $"未找到配置文件，请先部署：\n{string.Join("\n", config.MissingFiles())}\n\n" +
                "参考 sample_config 目录生成 db.ini（AES 加密连接串）与 config.ini。",
                "配置缺失", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        var loginWindow = new Views.LoginWindow();
        loginWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            var db = ServiceProvider?.GetService<DatabaseService>();
            db?.ClearPool();
        }
        catch { }

        try { Log.Information("ExcelMaker 退出"); Log.CloseAndFlush(); }
        catch { }

        base.OnExit(e);
    }
}

namespace ExcelMaker.ViewModels;

public partial class LoginViewModel : ObservableObject
{
    private readonly DatabaseService _db;
    private readonly CryptoService _crypto;
    private static readonly string LoginDataPath = Path.Combine(@"D:\ExcelMaker", "login.dat");

    [ObservableProperty] private string _username = string.Empty;
    [ObservableProperty] private string _password = string.Empty;
    [ObservableProperty] private string _errorMessage = string.Empty;
    [ObservableProperty] private bool _isLoggingIn;
    [ObservableProperty] private bool _rememberPassword;

    public LoginViewModel(DatabaseService db, CryptoService crypto)
    {
        _db = db;
        _crypto = crypto;
        LoadSavedLogin();
    }

    private void LoadSavedLogin()
    {
        try
        {
            if (!File.Exists(LoginDataPath)) return;
            var lines = File.ReadAllLines(LoginDataPath);
            if (lines.Length >= 3 && lines[0] == "1")
            {
                Username = _crypto.Decrypt(lines[1]);
                Password = _crypto.Decrypt(lines[2]);
                RememberPassword = true;
            }
            else
            {
                RememberPassword = false;
                File.Delete(LoginDataPath);
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "加载登录信息失败");
            RememberPassword = false;
        }
    }

    private void SaveLogin()
    {
        try
        {
            var dir = Path.GetDirectoryName(LoginDataPath)!;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            if (RememberPassword && !string.IsNullOrEmpty(Username))
                File.WriteAllLines(LoginDataPath, new[] { "1", _crypto.Encrypt(Username), _crypto.Encrypt(Password) });
            else
                ClearSavedLogin();
        }
        catch (Exception ex) { Log.Warning(ex, "保存登录信息失败"); }
    }

    public void ClearSavedLogin()
    {
        try { if (File.Exists(LoginDataPath)) File.Delete(LoginDataPath); }
        catch (Exception ex) { Log.Warning(ex, "清除登录信息失败"); }
    }

    [RelayCommand(CanExecute = nameof(CanLogin))]
    private async Task LoginAsync()
    {
        if (string.IsNullOrWhiteSpace(Username) || string.IsNullOrWhiteSpace(Password))
        {
            ErrorMessage = "请输入账号和密码";
            return;
        }

        IsLoggingIn = true;
        ErrorMessage = string.Empty;

        try
        {
            var md5Pwd = Md5Helper.Compute(Password);
            if (!await _db.ValidateLoginAsync(Username, md5Pwd))
            {
                ErrorMessage = "账号或密码错误";
                return;
            }

            var fullName = await _db.GetUserFullNameAsync(Username);
            SaveLogin();

            var mainVm = App.ServiceProvider.GetRequiredService<MainViewModel>();
            mainVm.Username = Username.Trim();   // 工号，用于 MERGE 的 AUTHOR
            mainVm.CurrentUser = fullName;

            var mainWindow = new Views.MainWindow { DataContext = mainVm };
            mainWindow.Show();

            Application.Current.Windows
                .OfType<Window>()
                .FirstOrDefault(w => w.DataContext == this)
                ?.Close();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "登录失败");
            ErrorMessage = $"网络异常，请重试: {ex.Message}";
        }
        finally { IsLoggingIn = false; }
    }

    private bool CanLogin() => !IsLoggingIn;
}

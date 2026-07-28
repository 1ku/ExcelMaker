namespace ExcelMaker.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        // 主窗口关闭即退出程序
        Application.Current.Shutdown();
    }
}

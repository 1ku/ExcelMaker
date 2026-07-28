namespace ExcelMaker.Views;

public partial class LoginWindow : Window
{
    public LoginWindow()
    {
        InitializeComponent();
        DataContext = App.ServiceProvider.GetRequiredService<LoginViewModel>();

        PasswordBox.PasswordChanged += (s, e) =>
        {
            if (DataContext is LoginViewModel vm)
                vm.Password = PasswordBox.Password;
        };

        RememberCheckBox.Unchecked += (s, e) =>
        {
            if (DataContext is LoginViewModel vm)
                vm.ClearSavedLogin();
        };

        KeyDown += (s, e) =>
        {
            if (e.Key == Key.Enter && DataContext is LoginViewModel vm && vm.LoginCommand.CanExecute(null))
                vm.LoginCommand.Execute(null);
        };

        Loaded += (s, e) =>
        {
            if (DataContext is LoginViewModel vm && !string.IsNullOrEmpty(vm.Password))
                PasswordBox.Password = vm.Password;
            Keyboard.Focus(string.IsNullOrEmpty(((LoginViewModel)DataContext).Username) ? UsernameBox : PasswordBox);
        };
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState == MouseButtonState.Pressed) DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();
}

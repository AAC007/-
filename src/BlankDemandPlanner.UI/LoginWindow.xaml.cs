using System.Windows;
using System.Windows.Input;
using BlankDemandPlanner.UI.Services;

namespace BlankDemandPlanner.UI;

public partial class LoginWindow : Window
{
    private readonly IAppAuthService _authService;
    private bool _isFirstSetup;

    public LoginWindow(IAppAuthService authService)
    {
        InitializeComponent();
        _authService = authService;
        Loaded += LoginWindow_Loaded;
        PreviewKeyDown += LoginWindow_PreviewKeyDown;
    }

    private async void LoginWindow_Loaded(object sender, RoutedEventArgs e)
    {
        _isFirstSetup = !await _authService.HasAnyUsersAsync();
        if (_isFirstSetup)
        {
            ModeText.Text = "Первичная настройка администратора";
            SubmitButton.Content = "Применить";
            UserNameBox.Text = "admin";
            DisplayNameBox.Text = "Администратор";
            DisplayNameLabel.Visibility = Visibility.Visible;
            DisplayNameBox.Visibility = Visibility.Visible;
            ConfirmPasswordLabel.Visibility = Visibility.Visible;
            ConfirmPasswordBox.Visibility = Visibility.Visible;
        }
        else
        {
            var userNames = await _authService.GetLoginUserNamesAsync();
            UserNameBox.ItemsSource = userNames;
            UserNameBox.Text = await _authService.GetLastLoginUserNameAsync() ?? userNames.FirstOrDefault() ?? string.Empty;
            DisplayNameLabel.Visibility = Visibility.Collapsed;
            DisplayNameBox.Visibility = Visibility.Collapsed;
            ConfirmPasswordLabel.Visibility = Visibility.Collapsed;
            ConfirmPasswordBox.Visibility = Visibility.Collapsed;
        }

        UserNameBox.Focus();
    }

    private void LoginWindow_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter || !SubmitButton.IsEnabled)
        {
            return;
        }

        e.Handled = true;
        Submit_Click(SubmitButton, new RoutedEventArgs());
    }

    private async void Submit_Click(object sender, RoutedEventArgs e)
    {
        SubmitButton.IsEnabled = false;
        StatusText.Text = string.Empty;
        try
        {
            AuthResult result;
            if (_isFirstSetup)
            {
                if (!string.Equals(PasswordBox.Password, ConfirmPasswordBox.Password, StringComparison.Ordinal))
                {
                    StatusText.Text = "Пароли не совпадают.";
                    return;
                }

                result = await _authService.SetupFirstAdminAsync(UserNameBox.Text, DisplayNameBox.Text, PasswordBox.Password);
            }
            else
            {
                result = await _authService.LoginAsync(UserNameBox.Text, PasswordBox.Password);
            }

            if (result.IsSuccess)
            {
                DialogResult = true;
                Close();
                return;
            }

            StatusText.Text = result.Message;
        }
        finally
        {
            SubmitButton.IsEnabled = true;
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}

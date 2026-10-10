using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using LabSnapWord.Services;

namespace LabSnapWord.Views
{
    public partial class InstallerWindow : Window
    {
        private readonly bool _isStandaloneSetupMode;

        public InstallerWindow(bool isStandaloneSetupMode = false)
        {
            InitializeComponent();
            _isStandaloneSetupMode = isStandaloneSetupMode;
            InstallLocationText.Text = $"Install location: {AppInstallerService.GetDefaultInstallDirectory()}";

            if (!_isStandaloneSetupMode)
            {
                RunPortableButton.Content = "Cancel";
                LaunchAfterInstallCheckBox.Visibility = Visibility.Collapsed;
            }
        }

        private void RunPortableButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isStandaloneSetupMode)
            {
                MainWindow main = new MainWindow();
                System.Windows.Application.Current.MainWindow = main;
                main.Show();
            }
            Close();
        }

        private async void InstallNowButton_Click(object sender, RoutedEventArgs e)
        {
            bool createDesktop = DesktopShortcutCheckBox.IsChecked == true;
            bool createStartMenu = StartMenuCheckBox.IsChecked == true;
            bool launchAfter = LaunchAfterInstallCheckBox.IsChecked == true;

            InstallNowButton.IsEnabled = false;
            RunPortableButton.IsEnabled = false;
            InstallNowButton.Content = "Installing...";
            InstallProgressBar.Visibility = Visibility.Visible;

            try
            {
                string installedExe = await Task.Run(() =>
                    AppInstallerService.InstallApplication(createDesktop, createStartMenu));

                InstallProgressBar.IsIndeterminate = false;
                InstallProgressBar.Value = 100;
                InstallNowButton.Content = "Installed!";

                System.Windows.MessageBox.Show(
                    this,
                    "LabSnap Word has been installed on your PC!\n\n" +
                    "• Desktop Shortcut created\n" +
                    "• Start Menu Shortcut created\n\n" +
                    "You can now open LabSnap Word anytime directly from your Desktop without needing the USB drive.",
                    "Installation Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                if (_isStandaloneSetupMode && launchAfter)
                {
                    Process.Start(new ProcessStartInfo(installedExe)
                    {
                        UseShellExecute = true
                    });
                }

                Close();
            }
            catch (Exception ex)
            {
                InstallProgressBar.Visibility = Visibility.Collapsed;
                InstallNowButton.IsEnabled = true;
                RunPortableButton.IsEnabled = true;
                InstallNowButton.Content = "Install Now";

                System.Windows.MessageBox.Show(
                    this,
                    $"Installation failed:\n{ex.Message}",
                    "Installer Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}

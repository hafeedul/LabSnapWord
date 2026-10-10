using System;
using System.IO;
using System.Linq;
using System.Windows;
using LabSnapWord.Views;

namespace LabSnapWord
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            string exeName = Path.GetFileNameWithoutExtension(Environment.ProcessPath ?? string.Empty);
            bool isSetupExe = exeName.IndexOf("Setup", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              exeName.IndexOf("Install", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              e.Args.Any(a => string.Equals(a, "--install", StringComparison.OrdinalIgnoreCase));

            if (isSetupExe)
            {
                InstallerWindow installer = new InstallerWindow(isStandaloneSetupMode: true);
                MainWindow = installer;
                installer.Show();
            }
            else
            {
                MainWindow main = new MainWindow();
                MainWindow = main;
                main.Show();
            }
        }
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace LabSnapWord.Services
{
    public static class AppInstallerService
    {
        public static string GetDefaultInstallDirectory()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "Programs", "LabSnapWord");
        }

        public static string GetInstalledExePath()
        {
            return Path.Combine(GetDefaultInstallDirectory(), "LabSnapWord.exe");
        }

        public static bool IsRunningFromInstalledLocation()
        {
            string? currentExe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(currentExe))
            {
                return false;
            }

            string installedExe = GetInstalledExePath();
            return string.Equals(
                Path.GetFullPath(currentExe),
                Path.GetFullPath(installedExe),
                StringComparison.OrdinalIgnoreCase);
        }

        public static string InstallApplication(bool createDesktopShortcut, bool createStartMenuShortcut)
        {
            string? sourceExe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(sourceExe) || !File.Exists(sourceExe))
            {
                throw new InvalidOperationException("Cannot locate running executable package.");
            }

            string installDir = GetDefaultInstallDirectory();
            Directory.CreateDirectory(installDir);

            string targetExePath = Path.Combine(installDir, "LabSnapWord.exe");
            string iconPath = Path.Combine(installDir, "app_icon.ico");

            if (!string.Equals(Path.GetFullPath(sourceExe), Path.GetFullPath(targetExePath), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourceExe, targetExePath, overwrite: true);
            }

            var iconResource = System.Windows.Application.GetResourceStream(
                new Uri("pack://application:,,,/Assets/app_icon.ico"));
            if (iconResource != null)
            {
                using (Stream s = iconResource.Stream)
                using (FileStream fs = File.Create(iconPath))
                {
                    s.CopyTo(fs);
                }
            }

            if (createDesktopShortcut)
            {
                string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string desktopLnk = Path.Combine(desktopDir, "LabSnap Word.lnk");
                CreateWindowsShortcut(desktopLnk, targetExePath, installDir, iconPath);
            }

            if (createStartMenuShortcut)
            {
                string startMenuDir = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
                string startMenuLnk = Path.Combine(startMenuDir, "LabSnap Word.lnk");
                CreateWindowsShortcut(startMenuLnk, targetExePath, installDir, iconPath);
            }

            return targetExePath;
        }

        public static void CreateWindowsShortcut(string shortcutPath, string targetExePath, string workingDir, string iconPath)
        {
            string psScript = @"
$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut($env:LABSNAP_LNK_PATH)
$s.TargetPath = $env:LABSNAP_TARGET_EXE
$s.WorkingDirectory = $env:LABSNAP_WORK_DIR
$s.IconLocation = $env:LABSNAP_ICON_PATH
$s.Description = 'LabSnap Word - University Lab Screenshot Cropper & Word Report Builder'
$s.Save()
";
            string encoded = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(psScript));

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            psi.Environment["LABSNAP_LNK_PATH"] = shortcutPath;
            psi.Environment["LABSNAP_TARGET_EXE"] = targetExePath;
            psi.Environment["LABSNAP_WORK_DIR"] = workingDir;
            psi.Environment["LABSNAP_ICON_PATH"] = File.Exists(iconPath) ? iconPath : $"{targetExePath},0";

            using (Process? proc = Process.Start(psi))
            {
                proc?.WaitForExit(6000);
            }
        }
    }
}

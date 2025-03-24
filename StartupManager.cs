
using System;
using System.IO;
using IWshRuntimeLibrary;

class StartupManager
{
    public static void CreateStartupShortcut()
    {
        string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        string shortcutPath = Path.Combine(startupPath, "WindowsFormsBLEserver.lnk");
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        if (!File.Exists(shortcutPath))
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = exePath;
            shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
            shortcut.Save();
        }
    }
}

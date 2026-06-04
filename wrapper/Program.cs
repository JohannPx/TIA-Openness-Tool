using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace TiaOpennessTool;

static class Program
{
    const string AppName = "TIA Openness Tool";
    const string ExeName = "TiaOpennessTool.exe";
    const string GitHubApiUrl = "https://api.github.com/repos/JohannPx/TIA-Openness-Tool/releases/latest";
    const string ResourceName = "TiaOpennessTool.TIA-Openness-Tool_latest.ps1";

    static readonly string InstallDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TiaOpennessTool");
    static readonly string VersionFile = Path.Combine(InstallDir, "version.json");

    [STAThread]
    static async Task<int> Main(string[] args)
    {
        try
        {
            if (InstallIfNeeded()) return 0;
            await UpdateSilently();
            LaunchScript();
        }
        catch
        {
            // Fallback: just launch the script regardless of errors
            try { LaunchScript(); } catch { }
        }
        return 0;
    }

    /// <summary>
    /// First launch: copy exe to AppData, create shortcuts, relaunch from there.
    /// </summary>
    static bool InstallIfNeeded()
    {
        var currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (currentExe == null || !currentExe.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            return false;

        var installedExe = Path.Combine(InstallDir, ExeName);

        // Already running from install dir
        if (string.Equals(Path.GetFullPath(currentExe), Path.GetFullPath(installedExe), StringComparison.OrdinalIgnoreCase))
            return false;

        Directory.CreateDirectory(InstallDir);
        File.Copy(currentExe, installedExe, true);

        // Create shortcuts via PowerShell (avoids COM interop / trimming issues)
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        CreateShortcut(Path.Combine(desktop, $"{AppName}.lnk"), installedExe);

        var startMenu = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs");
        Directory.CreateDirectory(startMenu);
        CreateShortcut(Path.Combine(startMenu, $"{AppName}.lnk"), installedExe);

        // Initialize version.json
        if (!File.Exists(VersionFile))
            File.WriteAllText(VersionFile, """{"version":"0.0.0"}""");

        // Relaunch from install dir and exit
        Process.Start(new ProcessStartInfo(installedExe) { UseShellExecute = true });
        return true;
    }

    static void CreateShortcut(string lnkPath, string targetPath)
    {
        var dir = Path.GetDirectoryName(lnkPath);
        if (dir != null && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        var escapedLnk = lnkPath.Replace("'", "''");
        var escapedTarget = targetPath.Replace("'", "''");
        var escapedDir = Path.GetDirectoryName(targetPath)?.Replace("'", "''") ?? "";

        var ps = $"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{escapedLnk}');" +
                 $"$s.TargetPath='{escapedTarget}';" +
                 $"$s.WorkingDirectory='{escapedDir}';" +
                 $"$s.Description='{AppName}';" +
                 "$s.Save()";

        var psi = new ProcessStartInfo("powershell.exe", $"-NoProfile -Command \"{ps}\"")
        {
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            UseShellExecute = false
        };
        Process.Start(psi)?.WaitForExit(5000);
    }

    /// <summary>
    /// Check GitHub Releases for a newer version; download and swap if found.
    /// </summary>
    static async Task UpdateSilently()
    {
        var currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        var installedExe = Path.Combine(InstallDir, ExeName);
        if (currentExe == null ||
            !string.Equals(Path.GetFullPath(currentExe), Path.GetFullPath(installedExe), StringComparison.OrdinalIgnoreCase))
            return;

        var localVersion = "0.0.0";
        if (File.Exists(VersionFile))
        {
            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(VersionFile));
                localVersion = doc.RootElement.GetProperty("version").GetString() ?? "0.0.0";
            }
            catch { /* corrupted file, treat as 0.0.0 */ }
        }

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        http.DefaultRequestHeaders.Add("User-Agent", "TiaOpennessTool");

        var json = await http.GetStringAsync(GitHubApiUrl);
        using var release = JsonDocument.Parse(json);
        var root = release.RootElement;

        var remoteTag = root.GetProperty("tag_name").GetString() ?? "";
        var remoteVersion = remoteTag.TrimStart('v');

        if (remoteVersion == localVersion) return;

        // Find the .exe asset
        string? downloadUrl = null;
        foreach (var asset in root.GetProperty("assets").EnumerateArray())
        {
            var name = asset.GetProperty("name").GetString() ?? "";
            if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                downloadUrl = asset.GetProperty("browser_download_url").GetString();
                break;
            }
        }
        if (downloadUrl == null) return;

        // Download new exe to temp
        var tempExe = Path.Combine(Path.GetTempPath(), "TiaOpennessTool_update.exe");
        var bytes = await http.GetByteArrayAsync(downloadUrl);
        await File.WriteAllBytesAsync(tempExe, bytes);

        // Save new version
        File.WriteAllText(VersionFile,
            $$$"""{"version":"{{{remoteVersion}}}","date":"{{{DateTime.Now:yyyy-MM-dd}}}"}""");

        // Write batch to replace exe and relaunch
        var batchPath = Path.Combine(Path.GetTempPath(), "tia_openness_update.cmd");
        File.WriteAllText(batchPath,
            $"""
            @echo off
            timeout /t 2 /nobreak >nul
            copy /y "{tempExe}" "{installedExe}" >nul
            start "" "{installedExe}"
            del "{tempExe}" >nul 2>&1
            del "%~f0" >nul 2>&1
            """, Encoding.ASCII);

        Process.Start(new ProcessStartInfo("cmd.exe", $"""/c "{batchPath}" """)
        {
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            UseShellExecute = false
        });

        Environment.Exit(0);
    }

    /// <summary>
    /// Extract the embedded .ps1 script and run it via powershell.exe (STA, required by WPF).
    /// </summary>
    static void LaunchScript()
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), $"TIA-Openness-Tool_{Guid.NewGuid():N}.ps1");

        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
        {
            if (stream == null)
                throw new InvalidOperationException($"Embedded resource '{ResourceName}' not found.");
            using var fs = File.Create(scriptPath);
            stream.CopyTo(fs);
        }

        try
        {
            var psi = new ProcessStartInfo("powershell.exe",
                $"""-Sta -WindowStyle Hidden -ExecutionPolicy Bypass -File "{scriptPath}" """)
            {
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false
            };

            var proc = Process.Start(psi);
            proc?.WaitForExit();
        }
        finally
        {
            try { File.Delete(scriptPath); } catch { }
        }
    }
}

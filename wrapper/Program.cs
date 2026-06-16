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

    // Variable d'environnement transmise au script PowerShell quand une mise à jour est
    // disponible mais que son téléchargement a échoué : l'app affiche alors un bandeau.
    const string UpdateNoticeEnvVar = "TIA_OPENNESS_UPDATE_AVAILABLE";

    static readonly string InstallDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TiaOpennessTool");
    static readonly string VersionFile = Path.Combine(InstallDir, "version.json");

    [STAThread]
    static async Task<int> Main(string[] args)
    {
        string? updateNotice = null;
        try
        {
            if (InstallIfNeeded()) return 0;
            updateNotice = await CheckAndUpdate();
        }
        catch
        {
            // Une erreur de mise à jour ne doit jamais empêcher le lancement.
        }
        // Fallback: lance le script quoi qu'il arrive
        try { LaunchScript(updateNotice); } catch { }
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
    /// Vérifie les Releases GitHub ; télécharge et applique la nouvelle version si elle existe.
    /// Retourne la version distante si une mise à jour est disponible mais que son
    /// téléchargement a échoué (à signaler dans l'app), sinon null.
    /// </summary>
    static async Task<string?> CheckAndUpdate()
    {
        var currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        var installedExe = Path.Combine(InstallDir, ExeName);
        if (currentExe == null ||
            !string.Equals(Path.GetFullPath(currentExe), Path.GetFullPath(installedExe), StringComparison.OrdinalIgnoreCase))
            return null;

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

        // 1. Vérifier la dernière version publiée (requête légère, timeout court).
        //    Si la machine est hors ligne, on sort sans rien signaler : pas de nag inutile.
        string remoteVersion;
        string? downloadUrl;
        try
        {
            using var api = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            api.DefaultRequestHeaders.Add("User-Agent", "TiaOpennessTool");
            var json = await api.GetStringAsync(GitHubApiUrl);
            using var release = JsonDocument.Parse(json);
            var root = release.RootElement;
            remoteVersion = (root.GetProperty("tag_name").GetString() ?? "").TrimStart('v');
            downloadUrl = FindExeAssetUrl(root);
        }
        catch
        {
            return null;
        }

        if (string.IsNullOrEmpty(remoteVersion) || remoteVersion == localVersion || downloadUrl == null)
            return null;

        // 2. Une mise à jour existe : télécharger l'exe (volumineux → timeout large, contrairement
        //    aux 10 s de la requête API : un lien lent ne doit pas faire échouer la MAJ).
        try
        {
            using var dl = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
            dl.DefaultRequestHeaders.Add("User-Agent", "TiaOpennessTool");
            var bytes = await dl.GetByteArrayAsync(downloadUrl);

            var tempExe = Path.Combine(Path.GetTempPath(), "TiaOpennessTool_update.exe");
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
            return null; // inatteignable
        }
        catch
        {
            // Téléchargement/échange échoué : on NE met PAS à jour version.json (nouvelle
            // tentative au prochain lancement) et on signale la version dispo à l'app.
            return remoteVersion;
        }
    }

    /// <summary>
    /// Retourne l'URL de téléchargement du premier asset .exe de la release (nom versionné
    /// ou non : la correspondance se fait sur l'extension, pas sur un nom figé).
    /// </summary>
    static string? FindExeAssetUrl(JsonElement releaseRoot)
    {
        foreach (var asset in releaseRoot.GetProperty("assets").EnumerateArray())
        {
            var name = asset.GetProperty("name").GetString() ?? "";
            if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                return asset.GetProperty("browser_download_url").GetString();
        }
        return null;
    }

    /// <summary>
    /// Extract the embedded .ps1 script and run it via powershell.exe (STA, required by WPF).
    /// </summary>
    static void LaunchScript(string? updateNotice = null)
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

            // Signale à l'app qu'une mise à jour est dispo mais n'a pas pu être téléchargée.
            if (!string.IsNullOrEmpty(updateNotice))
                psi.Environment[UpdateNoticeEnvVar] = updateNotice;

            var proc = Process.Start(psi);
            proc?.WaitForExit();
        }
        finally
        {
            try { File.Delete(scriptPath); } catch { }
        }
    }
}

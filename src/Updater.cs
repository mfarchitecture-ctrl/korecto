using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CorrecteurMail {
    internal sealed class ReleaseInfo {
        internal Version Version;
        internal string Tag, AssetUrl, Digest;
        internal long Size;
    }
    // Updates come from the GitHub Releases of this project only: one HTTPS request to read the latest
    // release, one to download Korecto.exe. Nothing about the user or the PC is sent.
    internal static class Updater {
        internal const string LatestUrl = "https://api.github.com/repos/mfarchitecture-ctrl/korecto/releases/latest";
        internal const string DownloadPrefix = "https://github.com/mfarchitecture-ctrl/korecto/releases/download/";
        internal const long MaxBytes = 40L * 1024 * 1024;
        internal static string UpdateFolder { get { return Path.Combine(AppFolders.Data, "maj"); } }

        internal static Version ParseVersion(string text) {
            if (String.IsNullOrWhiteSpace(text)) return null;
            text = text.Trim(); if (text.StartsWith("v", StringComparison.OrdinalIgnoreCase)) text = text.Substring(1);
            Version value; return Version.TryParse(text, out value) && value.Major >= 0 && value.Minor >= 0 ? value : null;
        }
        internal static bool IsNewer(string current, Version latest) {
            Version now = ParseVersion(current);
            return now != null && latest != null && Normalize(latest) > Normalize(now);
        }
        private static Version Normalize(Version v) { return new Version(v.Major, v.Minor, Math.Max(v.Build, 0)); }

        // Returns null when the answer is not a usable release (never throws on bad content).
        internal static ReleaseInfo ParseRelease(string json) {
            try {
                var root = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                if (root == null || !root.ContainsKey("tag_name") || !(root["tag_name"] is string)) return null;
                if (root.ContainsKey("draft") && root["draft"] is bool && (bool)root["draft"]) return null;
                if (root.ContainsKey("prerelease") && root["prerelease"] is bool && (bool)root["prerelease"]) return null;
                var info = new ReleaseInfo { Tag = (string)root["tag_name"] };
                info.Version = ParseVersion(info.Tag);
                if (info.Version == null || !root.ContainsKey("assets")) return null;
                foreach (object asset in (IEnumerable)root["assets"]) {
                    var item = asset as Dictionary<string, object>;
                    if (item == null || !item.ContainsKey("name") || !(item["name"] is string) || (string)item["name"] != SelfInstaller.ExeName) continue;
                    if (!item.ContainsKey("browser_download_url") || !(item["browser_download_url"] is string)) continue;
                    string url = (string)item["browser_download_url"];
                    if (!url.StartsWith(DownloadPrefix, StringComparison.Ordinal)) return null;
                    info.AssetUrl = url;
                    if (item.ContainsKey("size") && (item["size"] is int || item["size"] is long)) info.Size = Convert.ToInt64(item["size"]);
                    if (item.ContainsKey("digest") && item["digest"] is string) info.Digest = (string)item["digest"];
                    return info;
                }
                return null;
            } catch (ArgumentException) { return null; } catch (InvalidOperationException) { return null; } catch (InvalidCastException) { return null; }
        }

        // The file must be the announced size, a Windows program, and match the published SHA-256 when GitHub gives one.
        internal static bool Verify(string file, ReleaseInfo release) {
            var info = new FileInfo(file);
            if (!info.Exists || info.Length < 1024 || info.Length > MaxBytes) return false;
            if (release.Size > 0 && info.Length != release.Size) return false;
            using (var stream = File.OpenRead(file)) { if (stream.ReadByte() != 'M' || stream.ReadByte() != 'Z') return false; }
            if (String.IsNullOrEmpty(release.Digest)) return true;
            if (!release.Digest.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase)) return false;
            string expected = release.Digest.Substring(7).Trim().ToLowerInvariant();
            using (var sha = SHA256.Create()) using (var stream = File.OpenRead(file))
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() == expected;
        }
        // A release whose program announces an older version would be offered again after every update.
        internal static bool AnnouncesVersion(string file, ReleaseInfo release) {
            try {
                Version inside = ParseVersion(FileVersionInfo.GetVersionInfo(file).FileVersion);
                return inside != null && Normalize(inside) >= Normalize(release.Version);
            } catch (FileNotFoundException) { return false; }
        }
        internal static bool IsKnownError(Exception ex) {
            return ex is WebException || ex is IOException || ex is InvalidOperationException || ex is UnauthorizedAccessException || ex is TimeoutException;
        }

        internal static string Fetch(string url) {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = "Korecto/" + AppBrand.Version; request.Accept = "application/vnd.github+json"; request.Timeout = 10000; request.ReadWriteTimeout = 10000;
            using (var response = (HttpWebResponse)request.GetResponse()) using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8)) return reader.ReadToEnd();
        }
        internal static void Download(string url, string target) {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = "Korecto/" + AppBrand.Version; request.Timeout = 15000; request.ReadWriteTimeout = 30000;
            using (var response = (HttpWebResponse)request.GetResponse()) {
                if (response.ContentLength > MaxBytes) throw new InvalidOperationException("Fichier annoncé trop volumineux.");
                using (var input = response.GetResponseStream()) using (var output = File.Create(target)) {
                    var buffer = new byte[81920]; long total = 0; int read;
                    while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
                        total += read; if (total > MaxBytes) throw new InvalidOperationException("Fichier téléchargé trop volumineux.");
                        output.Write(buffer, 0, read);
                    }
                }
            }
        }
        // Null when Korecto is up to date.
        internal static Task<ReleaseInfo> CheckAsync(Func<string, string> fetch = null) {
            return Task.Run(() => {
                var release = ParseRelease((fetch ?? Fetch)(LatestUrl));
                return release != null && IsNewer(AppBrand.Version, release.Version) ? release : null;
            });
        }
        internal static Task<string> DownloadForUpdate(ReleaseInfo release, Action<string, string> download = null, string folder = null) {
            return Task.Run(() => {
                folder = folder ?? UpdateFolder;
                if (Directory.Exists(folder)) Directory.Delete(folder, true);
                Directory.CreateDirectory(folder);
                string target = Path.Combine(folder, SelfInstaller.ExeName);
                (download ?? Download)(release.AssetUrl, target);
                if (!Verify(target, release)) throw new InvalidOperationException("Le fichier téléchargé est incomplet ou ne correspond pas à la version publiée.");
                if (!AnnouncesVersion(target, release)) throw new InvalidOperationException("La version publiée ne correspond pas à son numéro. Mise à jour annulée.");
                return target;
            });
        }
        internal static bool WaitUntil(Func<bool> done, int timeoutMs, int pollMs) {
            var watch = Stopwatch.StartNew();
            while (!done()) { if (watch.ElapsedMilliseconds >= timeoutMs) return false; Thread.Sleep(pollMs); }
            return true;
        }
        // Runs from the downloaded copy (--appliquer-maj): waits for the old instance to quit, replaces the installed exe, restarts it.
        internal static void ApplyAndRestart(string downloaded, SelfInstaller installer = null, Func<bool> running = null, int waitMs = 30000) {
            installer = installer ?? new SelfInstaller();
            if (!WaitUntil(() => !(running ?? SelfInstaller.IsRunning)(), waitMs, 250)) throw new InvalidOperationException("Korecto ne s’est pas fermé : mise à jour annulée.");
            for (int attempt = 0; ; attempt++) {
                try { installer.Install(downloaded); break; }
                catch (IOException) { if (attempt >= 12) throw; Thread.Sleep(500); }
                catch (UnauthorizedAccessException) { if (attempt >= 12) throw; Thread.Sleep(500); }
            }
        }
        internal static void CleanOldDownloads() {
            try { if (Directory.Exists(UpdateFolder)) Directory.Delete(UpdateFolder, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }
    internal sealed class UpdatePreference {
        private readonly string path;
        internal UpdatePreference(string folder = null) { path = Path.Combine(folder ?? AppFolders.Data, "maj-desactivee.txt"); }
        internal bool Enabled { get { return !File.Exists(path); } }
        internal void Save(bool enabled) {
            if (enabled) { if (File.Exists(path)) File.Delete(path); return; }
            Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "1");
        }
    }
    internal static class UpdateUi {
        internal static bool IsInstalledCopy() { return new SelfInstaller().IsInstalledCopy(Application.ExecutablePath); }
        // Manual check from Paramètres.
        internal static async Task Interactive(IWin32Window owner, Action exitApplication) {
            if (!IsInstalledCopy()) { MessageBox.Show(owner, "Cette copie de Korecto n’est pas installée. Double-cliquez sur Korecto.exe pour l’installer : les mises à jour seront alors possibles.", "Mise à jour de Korecto"); return; }
            ReleaseInfo release;
            try { release = await Updater.CheckAsync(); }
            catch (Exception ex) {
                if (!Updater.IsKnownError(ex)) throw;
                MessageBox.Show(owner, "Impossible de joindre GitHub pour le moment. Vérifiez votre connexion Internet et réessayez.", "Mise à jour de Korecto"); return;
            }
            if (release == null) { MessageBox.Show(owner, "Korecto est à jour (version " + AppBrand.Version + ").", "Mise à jour de Korecto"); return; }
            await Offer(owner, release, exitApplication);
        }
        internal static async Task Offer(IWin32Window owner, ReleaseInfo release, Action exitApplication) {
            string question = "Korecto " + release.Version.ToString(3) + " est disponible (vous avez la version " + AppBrand.Version + ").\n\nMettre à jour maintenant ? Korecto se ferme, se met à jour et se relance. Vos réglages sont conservés.";
            if (MessageBox.Show(owner, question, "Mise à jour de Korecto", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try {
                string file = await Updater.DownloadForUpdate(release);
                Process.Start(new ProcessStartInfo(file, "--appliquer-maj") { UseShellExecute = false });
                exitApplication();
            } catch (Exception ex) {
                if (!(Updater.IsKnownError(ex) || ex is System.ComponentModel.Win32Exception)) throw;
                MessageBox.Show(owner, "Mise à jour impossible : " + ex.Message, "Mise à jour de Korecto");
            }
        }
        // Run by the downloaded copy: no window unless something goes wrong.
        internal static void RunApply() {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try {
                var installer = new SelfInstaller();
                Updater.ApplyAndRestart(Application.ExecutablePath, installer);
                Process.Start(installer.InstalledExe);
            } catch (Exception ex) {
                if (!(Updater.IsKnownError(ex) || ex is System.ComponentModel.Win32Exception || ex is System.Runtime.InteropServices.COMException)) throw;
                MessageBox.Show("Mise à jour impossible : " + ex.Message + "\n\nVous pouvez relancer Korecto normalement ; l’ancienne version reste installée.", "Mise à jour de Korecto");
            }
        }
    }
}

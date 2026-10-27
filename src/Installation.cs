using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CorrecteurMail {
    // Korecto.exe installs itself: nothing else to download, no script, no administrator rights.
    internal interface IInstallRegistry {
        string ReadRun();
        void WriteRun(string command);
        void RemoveRun();
        void WriteUninstall(Dictionary<string, object> values);
        void RemoveUninstall();
    }
    internal sealed class WindowsInstallRegistry : IInstallRegistry {
        private const string RunPath = @"Software\Microsoft\Windows\CurrentVersion\Run", RunName = "MFArchiCorrecteurMail";
        private const string UninstallPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\Korecto";
        public string ReadRun() { using (var key = Registry.CurrentUser.OpenSubKey(RunPath)) return key == null ? null : key.GetValue(RunName) as string; }
        public void WriteRun(string command) { using (var key = Registry.CurrentUser.CreateSubKey(RunPath)) key.SetValue(RunName, command, RegistryValueKind.String); }
        public void RemoveRun() { using (var key = Registry.CurrentUser.OpenSubKey(RunPath, true)) if (key != null) key.DeleteValue(RunName, false); }
        public void WriteUninstall(Dictionary<string, object> values) {
            using (var key = Registry.CurrentUser.CreateSubKey(UninstallPath)) foreach (var pair in values) key.SetValue(pair.Key, pair.Value, pair.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
        }
        public void RemoveUninstall() { Registry.CurrentUser.DeleteSubKeyTree(UninstallPath, false); }
    }
    internal sealed class SelfInstaller {
        internal const string ExeName = "Korecto.exe", MutexName = "Local\\CorrecteurMailChatGPT", DeclinedFile = "sans-installation.txt";
        internal readonly string ProgramsFolder, DataFolder, MenuFolder;
        private readonly IInstallRegistry registry;
        private readonly Action<string> deleteFolderLater;
        internal SelfInstaller(string programs = null, string data = null, string menu = null, IInstallRegistry installRegistry = null, Action<string> deleteLater = null) {
            ProgramsFolder = Path.GetFullPath(programs ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Korecto"));
            DataFolder = Path.GetFullPath(data ?? AppFolders.Data);
            MenuFolder = Path.GetFullPath(menu ?? Environment.GetFolderPath(Environment.SpecialFolder.Programs));
            registry = installRegistry ?? new WindowsInstallRegistry();
            deleteFolderLater = deleteLater ?? DeleteAfterExit;
        }
        internal string InstalledExe { get { return Path.Combine(ProgramsFolder, ExeName); } }
        internal bool Installed { get { return File.Exists(InstalledExe); } }
        internal bool IsInstalledCopy(string exe) { return String.Equals(Path.GetDirectoryName(Path.GetFullPath(exe)), ProgramsFolder, StringComparison.OrdinalIgnoreCase); }
        internal static string Hash(string file) {
            using (var sha = SHA256.Create()) using (var stream = File.OpenRead(file)) return BitConverter.ToString(sha.ComputeHash(stream));
        }
        internal static bool IsRunning() { Mutex existing; if (!Mutex.TryOpenExisting(MutexName, out existing)) return false; existing.Dispose(); return true; }
        private string ShortcutPath { get { return Path.Combine(MenuFolder, "Korecto.lnk"); } }
        internal void Install(string sourceExe) {
            Directory.CreateDirectory(ProgramsFolder);
            File.Copy(sourceExe, InstalledExe, true);
            foreach (string old in new [] { "DESINSTALLER.ps1", "DESINSTALLER.cmd" }) { string path = Path.Combine(ProgramsFolder, old); if (File.Exists(path)) File.Delete(path); }
            // Settings of the portable version next to the downloaded exe: this PC's only, never overwritten.
            AppFolders.MigrateLegacy(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(sourceExe)), "PARAMETRES_IA"), DataFolder);
            Directory.CreateDirectory(MenuFolder);
            CreateShortcut(ShortcutPath);
            // Windows startup: moved to the installed program only if it was already enabled.
            if (!String.IsNullOrWhiteSpace(registry.ReadRun())) registry.WriteRun(LocalCommand.Quote(InstalledExe) + " --startup");
            registry.WriteUninstall(new Dictionary<string, object> {
                { "DisplayName", "Korecto" }, { "Publisher", "Maxime Ferrand" }, { "InstallLocation", ProgramsFolder },
                { "DisplayIcon", InstalledExe }, { "UninstallString", LocalCommand.Quote(InstalledExe) + " --desinstaller" },
                { "NoModify", 1 }, { "NoRepair", 1 }
            });
        }
        internal void Uninstall(bool removeSettings) {
            if (File.Exists(ShortcutPath)) File.Delete(ShortcutPath);
            // The startup entry is removed only when it points to this installed program.
            string run = registry.ReadRun();
            if (!String.IsNullOrWhiteSpace(run) && run.IndexOf(InstalledExe, StringComparison.OrdinalIgnoreCase) >= 0) registry.RemoveRun();
            registry.RemoveUninstall();
            if (Directory.Exists(ProgramsFolder)) deleteFolderLater(ProgramsFolder);
            if (removeSettings && Directory.Exists(DataFolder)) Directory.Delete(DataFolder, true);
        }
        internal void RememberDeclined() { Directory.CreateDirectory(DataFolder); File.WriteAllText(Path.Combine(DataFolder, DeclinedFile), "1"); }
        internal bool Declined { get { return File.Exists(Path.Combine(DataFolder, DeclinedFile)); } }
        private void CreateShortcut(string path) {
            object shell = null, link = null;
            try {
                shell = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
                link = ((dynamic)shell).CreateShortcut(path);
                ((dynamic)link).TargetPath = InstalledExe; ((dynamic)link).WorkingDirectory = ProgramsFolder;
                ((dynamic)link).Description = "Korecto : correcteur de mails"; ((dynamic)link).Save();
            } finally { Native.Release(link); Native.Release(shell); }
        }
        private static void DeleteAfterExit(string folder) {
            // The uninstaller is the installed exe itself: wait for it to quit, then remove its folder.
            Process.Start(new ProcessStartInfo("cmd.exe", "/c ping 127.0.0.1 -n 3 >nul & rmdir /s /q " + LocalCommand.Quote(folder)) { CreateNoWindow = true, UseShellExecute = false });
        }
        // Double-click on a downloaded Korecto.exe: install it (or update it) instead of just running it.
        // Returns true when this process should stop (installed copy started, or cancelled).
        internal static bool OfferInstall(string exe) {
            var self = new SelfInstaller();
            if (Environment.GetEnvironmentVariable("KORECTO_SANS_INSTALLATION") != null || self.IsInstalledCopy(exe)) return false;
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try {
                if (self.Installed) {
                    if (Hash(exe) == Hash(self.InstalledExe)) {
                        if (!IsRunning()) Process.Start(self.InstalledExe);
                        else MessageBox.Show("Korecto est déjà installé et ouvert : retrouvez son icône près de l’horloge.", "Korecto");
                        return true;
                    }
                    var update = MessageBox.Show("Une autre version de Korecto est installée.\n\nMettre à jour ? Vos réglages sont conservés.\n\nNon : lancer cette copie sans installer.", "Korecto", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (update == DialogResult.Cancel) return true;
                    if (update == DialogResult.No) return false;
                } else {
                    if (self.Declined) return false;
                    var choice = MessageBox.Show("Installer Korecto sur cet ordinateur ?\n\nOui : installation pour votre session Windows (menu Démarrer), sans droits administrateur.\nNon : lancer sans installer, et ne plus demander.\nAnnuler : quitter.", "Korecto", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (choice == DialogResult.Cancel) return true;
                    if (choice == DialogResult.No) { try { self.RememberDeclined(); } catch (IOException) { } catch (UnauthorizedAccessException) { } return false; }
                }
                if (IsRunning()) { MessageBox.Show("Korecto est ouvert. Clic droit sur son icône près de l’horloge, puis Quitter, et relancez ce fichier.", "Korecto"); return true; }
                self.Install(exe);
                Process.Start(self.InstalledExe);
                return true;
            } catch (Exception ex) {
                if (!(ex is IOException || ex is UnauthorizedAccessException || ex is System.ComponentModel.Win32Exception || ex is COMException || ex is System.Security.SecurityException)) throw;
                MessageBox.Show("Installation impossible : " + ex.Message, "Korecto"); return true;
            }
        }
        // Windows "Applications" runs: Korecto.exe --desinstaller
        internal static void RunUninstallUi() {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            if (IsRunning()) { MessageBox.Show("Korecto est ouvert. Clic droit sur son icône près de l’horloge, puis Quitter, et relancez la désinstallation.", "Korecto"); return; }
            if (MessageBox.Show("Désinstaller Korecto ?", "Korecto", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            bool settings = MessageBox.Show("Supprimer aussi vos réglages (thème, choix de l’IA) ?", "Korecto", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            try { new SelfInstaller().Uninstall(settings); MessageBox.Show("Korecto est désinstallé.", "Korecto"); }
            catch (Exception ex) {
                if (!(ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException)) throw;
                MessageBox.Show("Désinstallation incomplète : " + ex.Message, "Korecto");
            }
        }
    }
}

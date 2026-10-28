using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CorrecteurMail {
    internal static class AppBrand {
        internal const string Name = "Korecto";
        internal const string Version = "1.0.2";
        internal static Icon LoadIcon(int size) { return LoadIcon(size, "App.ico"); }
        internal static Icon LoadNotificationIcon(int size) { return LoadIcon(size, "Notification.ico"); }
        private static Icon LoadIcon(int size, string resource) {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CorrecteurMail." + resource)) {
                if (stream == null) throw new InvalidOperationException("Icône de l’application manquante.");
                using (var icon = new Icon(stream, new Size(size, size))) return (Icon)icon.Clone();
            }
        }
        internal static Image LoadLogo() {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CorrecteurMail.Logo.png")) {
                if (stream == null) throw new InvalidOperationException("Logo de l’application manquant.");
                using (var image = Image.FromStream(stream)) return new Bitmap(image);
            }
        }
        internal static void SetIcon(Form form) {
            var icon = LoadIcon(32); form.Icon = icon; form.Disposed += (s,e) => icon.Dispose();
        }
    }
    internal interface IStartupStore {
        string Read();
        void Write(string command);
        void Remove();
    }
    internal sealed class RegistryStartupStore : IStartupStore {
        private const string KeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string ValueName = "MFArchiCorrecteurMail";
        public string Read() { using (var key = Registry.CurrentUser.OpenSubKey(KeyPath)) return key == null ? null : key.GetValue(ValueName) as string; }
        public void Write(string command) { using (var key = Registry.CurrentUser.CreateSubKey(KeyPath)) key.SetValue(ValueName, command, RegistryValueKind.String); }
        public void Remove() { using (var key = Registry.CurrentUser.OpenSubKey(KeyPath, true)) if (key != null) key.DeleteValue(ValueName, false); }
    }
    internal sealed class StartupPreference {
        private readonly IStartupStore store;
        private readonly string executable;
        internal StartupPreference(IStartupStore storage = null, string executablePath = null) { store = storage ?? new RegistryStartupStore(); executable = Path.GetFullPath(executablePath ?? Application.ExecutablePath); }
        internal bool Enabled { get { return !String.IsNullOrWhiteSpace(store.Read()); } }
        internal string Command { get { return LocalCommand.Quote(executable) + " --startup"; } }
        internal void Save(bool enabled) {
            string existing = store.Read();
            if (enabled) {
                if (Command.Length > 260) throw new InvalidOperationException("Chemin trop long pour le démarrage Windows. Placez le correcteur dans un dossier plus court.");
                if (!String.Equals(existing, Command, StringComparison.OrdinalIgnoreCase)) store.Write(Command);
            } else if (!String.IsNullOrWhiteSpace(existing)) store.Remove();
        }
    }
    internal sealed class SettingsDialog : Form {
        internal readonly CheckBox Startup = new CheckBox(), Updates = new CheckBox();
        internal readonly Button UpdateNow = new SoftButton();
        internal readonly Button Save = new SoftButton { Primary = true }, AiChoice = new SoftButton();
        internal readonly Button ThemeChoice;
        private readonly Label message = new Label(), aiSummary = new Label();
        private readonly StartupPreference startup;
        private readonly UpdatePreference updatePref;
        // AI chosen in its dialog, applied only by Enregistrer.
        internal AiSettings PendingAi { get; private set; }
        internal SettingsDialog(Func<AiSettings> currentAi, Func<Form, AiSettings, AiSettings> chooseAi, StartupPreference startupPreference = null, string appearancePath = null, Action<AiSettings> applyAi = null, Func<Form, Task> checkUpdates = null, UpdatePreference updatePreference = null) {
            SuspendLayout();
            startup = startupPreference ?? new StartupPreference();
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            Text = "Paramètres · " + AppBrand.Name; ClientSize = new Size(510, 480);
            Font = new Font("Segoe UI", 10); FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false; StartPosition = FormStartPosition.CenterParent;
            AppBrand.SetIcon(this);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(22), ColumnCount = 1, RowCount = 11 };
            foreach (int height in new [] { 38, 30, 36, 44, 36, 44, 30, 42, 42, 40, 44 }) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            layout.Controls.Add(new Label { Text = "Paramètres", Dock = DockStyle.Fill, Font = new Font("Segoe UI Semibold", 16) }, 0, 0);
            layout.Controls.Add(new Label { Text = "Général", Dock = DockStyle.Fill, ForeColor = Palette.Muted }, 0, 1);
            Startup.Text = "Lancer au démarrage de Windows"; Startup.Dock = DockStyle.Fill;
            try { Startup.Checked = startup.Enabled; }
            catch (Exception ex) { if (!IsSettingsError(ex)) throw; Startup.Enabled = false; message.Text = "Réglage de démarrage inaccessible sur cette session Windows."; }
            layout.Controls.Add(Startup, 0, 2);
            layout.Controls.Add(new Label { Text = "Démarrage discret dans la zone de notification à l’ouverture de votre session. Gardez l’exécutable au même emplacement.", Dock = DockStyle.Fill, ForeColor = Palette.Muted, Font = new Font("Segoe UI", 9) }, 0, 3);
            updatePref = updatePreference ?? new UpdatePreference();
            try { Updates.Checked = updatePref.Enabled; } catch (IOException) { Updates.Checked = true; }
            Updates.Text = "Rechercher les mises à jour au démarrage"; Updates.Dock = DockStyle.Fill; layout.Controls.Add(Updates, 0, 4);
            var versionRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            versionRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            versionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); versionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            var version = new LinkLabel { Text = "Version " + AppBrand.Version, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, LinkColor = Palette.Accent, ActiveLinkColor = Palette.Accent };
            version.LinkClicked += (s,e) => MessageBox.Show(this, AppAssets.Read(AppDomain.CurrentDomain.BaseDirectory, "NOUVEAUTES.txt"), "Nouveautés de Korecto");
            versionRow.Controls.Add(version, 0, 0);
            UpdateNow.Text = "Rechercher une mise à jour"; UpdateNow.Dock = DockStyle.Fill; UpdateNow.Margin = new Padding(3, 4, 0, 4); UpdateNow.Enabled = checkUpdates != null;
            UpdateNow.Click += async (s,e) => {
                UpdateNow.Enabled = false; message.Text = "Recherche en cours…";
                try { await checkUpdates(this); } finally { if (!IsDisposed) { UpdateNow.Enabled = true; message.Text = ""; } }
            };
            versionRow.Controls.Add(UpdateNow, 1, 0); layout.Controls.Add(versionRow, 0, 5);
            layout.Controls.Add(new Label { Text = "Apparence et intelligence artificielle", Dock = DockStyle.Fill, ForeColor = Palette.Muted }, 0, 6);
            var choices = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = Padding.Empty };
            choices.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 42)); choices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            ThemeChoice = Palette.CreateToggle(appearancePath, value => message.Text = value);
            AiChoice.Text = "Choisir l’IA…"; AiChoice.Dock = DockStyle.Fill;
            AiChoice.Click += (s,e) => {
                var chosen = chooseAi(this, PendingAi ?? currentAi());
                if (chosen == null) return;
                var current = currentAi();
                PendingAi = chosen.Provider == current.Provider && chosen.Model == current.Model ? null : chosen;
                aiSummary.Text = chosen.ProviderName + " · " + chosen.ModelName + (PendingAi != null ? " · appliqué après Enregistrer" : "");
            };
            choices.Controls.Add(ThemeChoice); choices.Controls.Add(AiChoice); layout.Controls.Add(choices, 0, 7);
            aiSummary.Text = currentAi().ProviderName + " · " + currentAi().ModelName; aiSummary.Dock = DockStyle.Fill; layout.Controls.Add(aiSummary, 0, 8);
            message.Dock = DockStyle.Fill; message.ForeColor = Palette.Muted; message.Font = new Font("Segoe UI", 9); layout.Controls.Add(message, 0, 9);
            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = Padding.Empty };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var cancel = new SoftButton { Text = "Annuler", DialogResult = DialogResult.Cancel, Dock = DockStyle.Fill };
            Save.Text = "Enregistrer"; Save.Dock = DockStyle.Fill;
            Save.Click += (s,e) => {
                try { if (Startup.Enabled) startup.Save(Startup.Checked); }
                catch (Exception ex) { if (!IsSettingsError(ex)) throw; message.Text = ex is InvalidOperationException ? ex.Message : "Windows a refusé l’enregistrement du démarrage automatique."; return; }
                try { updatePref.Save(Updates.Checked); }
                catch (Exception ex) { if (!(ex is IOException || ex is UnauthorizedAccessException)) throw; message.Text = "Réglage des mises à jour non enregistré · vérifiez l’accès au dossier."; return; }
                try { if (PendingAi != null && applyAi != null) applyAi(PendingAi); PendingAi = null; }
                catch (Exception ex) { if (!(ex is IOException || ex is UnauthorizedAccessException)) throw; message.Text = "Choix d’IA non enregistré · vérifiez l’accès au dossier de l’application."; return; }
                DialogResult = DialogResult.OK;
            };
            buttons.Controls.Add(cancel); buttons.Controls.Add(Save); layout.Controls.Add(buttons, 0, 10); Controls.Add(layout);
            HandleCreated += (s,e) => Palette.ApplyCaption(this);
            AcceptButton = Save; CancelButton = cancel; Palette.Apply(this);
            AutoScaleDimensions = new SizeF(96, 96); ResumeLayout(true);
        }
        private static bool IsSettingsError(Exception ex) { return ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is IOException || ex is InvalidOperationException; }
    }
}

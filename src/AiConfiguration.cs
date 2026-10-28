using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CorrecteurMail {
    // Settings live in the user's profile, never next to the program: it can sit in a
    // read-only folder (installed) or in a synced one (Dropbox locks files while syncing).
    internal static class AppFolders {
        internal static string Data {
            get {
                string forced = Environment.GetEnvironmentVariable("KORECTO_DONNEES");
                return Path.GetFullPath(String.IsNullOrWhiteSpace(forced) ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Korecto") : forced);
            }
        }
        // Earlier versions stored PARAMETRES_IA\<PC>.json and <PC>.theme.txt next to the exe.
        // Only this PC's files are copied, existing settings are never overwritten, originals stay.
        internal static int MigrateLegacy(string legacyFolder, string dataFolder, string machine = null) {
            int copied = 0;
            try {
                machine = String.Join("_", (machine ?? Environment.MachineName).Split(Path.GetInvalidFileNameChars()));
                foreach (var pair in new [] { new [] { machine + ".json", "ia.json" }, new [] { machine + ".theme.txt", "theme.txt" } }) {
                    string from = Path.Combine(legacyFolder, pair[0]), to = Path.Combine(dataFolder, pair[1]);
                    if (!File.Exists(from) || File.Exists(to) || new FileInfo(from).Length > 4096) continue;
                    Directory.CreateDirectory(dataFolder); File.Copy(from, to, false); copied++;
                }
            } catch (IOException) { } catch (UnauthorizedAccessException) { } catch (ArgumentException) { }
            return copied;
        }
    }
    internal static class AppAssets {
        internal static string Read(string directory, string name) {
            if (name != "CONSIGNES.txt" && name != "response.schema.json" && name != "NOUVEAUTES.txt") throw new ArgumentException("Ressource inconnue.");
            string local = Path.Combine(directory, name);
            if (File.Exists(local)) return File.ReadAllText(local, Encoding.UTF8);
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CorrecteurMail." + name)) {
                if (stream == null) throw new InvalidOperationException("Ressource manquante. Utilisez la version complète du correcteur.");
                using (var reader = new StreamReader(stream, Encoding.UTF8)) return reader.ReadToEnd();
            }
        }
        internal static string SchemaPath(string directory) {
            string local = Path.GetFullPath(Path.Combine(directory, "response.schema.json"));
            if (File.Exists(local)) return local;
            string content = Read(directory, "response.schema.json");
            string target = Path.Combine(AppFolders.Data, "RESSOURCES", "response.schema.json");
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            if (!File.Exists(target) || File.ReadAllText(target, Encoding.UTF8) != content) File.WriteAllText(target, content, new UTF8Encoding(false));
            return target;
        }
    }
    internal sealed class AiSettings {
        public AiSettings() { }
        public string Provider { get; set; }
        public string Model { get; set; }
        internal static AiSettings Default(string provider = "chatgpt") { return new AiSettings { Provider = provider, Model = provider == "claude" ? "haiku" : "gpt-6-luna" }; }
        internal bool Valid { get { return Provider == "chatgpt" && (Model == "gpt-6-luna" || Model == "gpt-6.1-terra" || Model == "gpt-6.1-sol") || Provider == "claude" && (Model == "haiku" || Model == "sonnet"); } }
        internal string ProviderName { get { return Provider == "claude" ? "Claude" : "ChatGPT"; } }
        internal string ModelName { get { return Model == "gpt-6-luna" ? "Luna" : Model == "gpt-6.1-terra" ? "Terra" : Model == "gpt-6.1-sol" ? "Sol" : Model == "sonnet" ? "Sonnet" : "Haiku"; } }
        internal AiSettings Copy() { return new AiSettings { Provider = Provider, Model = Model }; }
    }
    internal static class AiSettingsStore {
        internal static string DefaultPath {
            get { return Path.Combine(AppFolders.Data, "ia.json"); }
        }
        internal static AiSettings Load(string path = null) {
            try {
                path = path ?? DefaultPath;
                if (!File.Exists(path) || new FileInfo(path).Length > 4096) return null;
                var value = new JavaScriptSerializer().Deserialize<AiSettings>(File.ReadAllText(path, Encoding.UTF8));
                return value != null && value.Valid ? value : null;
            } catch (IOException) { return null; } catch (UnauthorizedAccessException) { return null; } catch (ArgumentException) { return null; } catch (InvalidOperationException) { return null; }
        }
        internal static void Save(AiSettings value, string path = null) {
            if (value == null || !value.Valid) throw new ArgumentException("Choix d’IA ou de modèle invalide.");
            path = Path.GetFullPath(path ?? DefaultPath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + ".tmp";
            try {
                File.WriteAllText(temporary, new JavaScriptSerializer().Serialize(value), new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            } finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
    internal static class AppearanceStore {
        internal static string DefaultPath { get { return Path.Combine(AppFolders.Data, "theme.txt"); } }
        internal static bool LoadDark(string path = null) {
            try { return File.ReadAllText(path ?? DefaultPath, Encoding.UTF8).Trim() != "clair"; }
            catch (IOException) { return true; } catch (UnauthorizedAccessException) { return true; }
        }
        internal static void Save(bool dark, string path = null) {
            path = Path.GetFullPath(path ?? DefaultPath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + ".tmp";
            try {
                File.WriteAllText(temporary, dark ? "sombre" : "clair", new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            } finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
    internal sealed class LocalCommand {
        internal string FileName, Prefix = "";
        internal ProcessStartInfo StartInfo(string arguments, string directory) {
            var info = new ProcessStartInfo(FileName, Prefix + arguments) {
                WorkingDirectory = directory, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
            };
            // Only the local account login is allowed. Never inherit a paid API or cloud provider.
            foreach (string name in new [] { "OPENAI_API_KEY", "CODEX_API_KEY", "OPENAI_BASE_URL", "ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL", "CLAUDE_CODE_OAUTH_TOKEN", "CLAUDE_CODE_USE_BEDROCK", "CLAUDE_CODE_USE_VERTEX", "CLAUDE_CODE_USE_FOUNDRY", "ANTHROPIC_FOUNDRY_API_KEY", "ANTHROPIC_MODEL", "ANTHROPIC_DEFAULT_HAIKU_MODEL", "ANTHROPIC_DEFAULT_SONNET_MODEL", "ANTHROPIC_DEFAULT_OPUS_MODEL" }) info.EnvironmentVariables.Remove(name);
            return info;
        }
        internal static string Quote(string value) {
            var result = new StringBuilder("\""); int slashes = 0;
            foreach (char c in value) {
                if (c == '\\') { slashes++; continue; }
                if (c == '"') result.Append('\\', slashes * 2 + 1); else result.Append('\\', slashes);
                slashes = 0; result.Append(c);
            }
            result.Append('\\', slashes * 2); return result.Append('"').ToString();
        }
    }
    internal sealed class AiConnection {
        internal string Provider, Message;
        internal bool Connected;
    }
    internal static class LocalAi {
        internal const string TokenNotice = "Chaque correction utilise les tokens de votre propre IA.";
        internal static bool ClaudeDesktopInstalled() {
            try {
                using (var packages = Registry.CurrentUser.OpenSubKey(@"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages")) {
                    if (packages != null) foreach (string name in packages.GetSubKeyNames()) if (name.StartsWith("Claude_", StringComparison.OrdinalIgnoreCase)) return true;
                }
            } catch (System.Security.SecurityException) { } catch (UnauthorizedAccessException) { } catch (IOException) { }
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Directory.Exists(Path.Combine(local, "AnthropicClaude")) || File.Exists(Path.Combine(local, "Programs", "Claude", "Claude.exe"));
        }
        private static string OnPath(string name) {
            foreach (string folder in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';')) {
                try { string path = Path.Combine(folder.Trim('"'), name); if (File.Exists(path)) return path; } catch (ArgumentException) { }
            } return null;
        }
        internal static LocalCommand Find(string provider) {
            if (provider == "chatgpt") { string path = AccountRunner.FindCodex(); return path == null ? null : new LocalCommand { FileName = path }; }
            string user = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            foreach (string path in new [] { OnPath("claude.exe"), Path.Combine(user, ".local", "bin", "claude.exe"), Path.Combine(user, ".claude", "local", "claude.exe") })
                if (path != null && File.Exists(path)) return new LocalCommand { FileName = path };
            string node = OnPath("node.exe");
            if (node != null) {
                var folders = new List<string> { Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm"), Path.GetDirectoryName(node) };
                folders.AddRange((Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'));
                foreach (string folder in folders) {
                    try {
                        string script = Path.Combine(folder.Trim('"'), "node_modules", "@anthropic-ai", "claude-code", "cli.js");
                        if (File.Exists(script)) return new LocalCommand { FileName = node, Prefix = LocalCommand.Quote(script) + " " };
                    } catch (ArgumentException) { }
                }
            }
            return null;
        }
        internal static bool IsAccountLogin(string provider, int code, string output) {
            if (code != 0) return false;
            if (provider == "chatgpt") return output.IndexOf("logged in using chatgpt", StringComparison.OrdinalIgnoreCase) >= 0;
            try {
                var value = new JavaScriptSerializer().DeserializeObject(output) as Dictionary<string, object>;
                return value != null && value.ContainsKey("loggedIn") && value["loggedIn"] is bool && (bool)value["loggedIn"] && value.ContainsKey("authMethod") && value["authMethod"] as string == "claude.ai";
            } catch (ArgumentException) { return false; } catch (InvalidOperationException) { return false; }
        }
        internal static AiConnection Check(string provider, LocalCommand command = null, string directory = null) {
            string name = provider == "claude" ? "Claude Code" : "Codex";
            command = command ?? Find(provider);
            if (command == null) return new AiConnection {
                Provider = provider,
                Message = provider == "claude"
                    ? (ClaudeDesktopInstalled() ? "Claude Desktop installé ; client Claude Code (CLI) introuvable." : "Client Claude Code (CLI) introuvable sur cet ordinateur.")
                    : name + " non installé sur cet ordinateur."
            };
            try {
                using (var process = new Process { StartInfo = command.StartInfo(provider == "claude" ? "--safe-mode auth status" : "login status", directory ?? AppDomain.CurrentDomain.BaseDirectory) }) {
                    process.Start(); process.StandardInput.Close();
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(6000)) { try { process.Kill(); } catch { } return new AiConnection { Provider = provider, Message = "Connexion " + name + " non vérifiable · réessayez." }; }
                    string text = output.GetAwaiter().GetResult();
                    if (provider == "chatgpt") text += "\n" + error.GetAwaiter().GetResult();
                    else error.GetAwaiter().GetResult();
                    bool connected = IsAccountLogin(provider, process.ExitCode, text);
                    return new AiConnection { Provider = provider, Connected = connected, Message = connected ? "Compte " + (provider == "claude" ? "Claude" : "ChatGPT") + " connecté sur cet ordinateur." : "Connectez votre compte dans " + name + ", puis Actualiser." };
                }
            } catch (Exception ex) {
                if (!(ex is IOException || ex is System.ComponentModel.Win32Exception || ex is InvalidOperationException || ex is UnauthorizedAccessException)) throw;
                return new AiConnection { Provider = provider, Message = "Connexion " + name + " inaccessible · vérifiez l’application locale." };
            }
        }
        internal static Task<AiConnection[]> Discover() {
            return Task.Run(() => new [] { Check("chatgpt"), Check("claude") });
        }
    }
    internal sealed class AiSettingsDialog : Form {
        internal readonly RadioButton Chatgpt = new RadioButton(), Claude = new RadioButton();
        internal readonly ComboBox Models = new ComboBox();
        internal readonly Button Accept = new SoftButton { Primary = true }, RefreshButton = new SoftButton();
        internal readonly Button ThemeChoice;
        private readonly Label chatStatus = new Label(), claudeStatus = new Label(), modelHelp = new Label();
        private readonly string settingsPath;
        private readonly AiSettings initial;
        private readonly Func<Task<AiConnection[]>> discover;
        private AiConnection[] connections = new AiConnection[0];
        private bool checking;
        private readonly bool persist;
        internal AiSettings Selected { get; private set; }
        internal AiSettingsDialog(AiSettings current = null, string path = null, Func<Task<AiConnection[]>> connectionDiscovery = null, bool autoCheck = true, string appearancePath = null, bool persist = true) {
            SuspendLayout();
            initial = current; settingsPath = path; discover = connectionDiscovery ?? LocalAi.Discover; this.persist = persist;
            Text = (current == null ? "Choisissez votre IA" : "Paramètres d’IA") + " · " + AppBrand.Name;
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(510, 420); MinimumSize = SizeFromClientSize(new Size(510, 420));
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen; BackColor = Palette.Background; ForeColor = Palette.Text; Font = new Font("Segoe UI", 10);
            AppBrand.SetIcon(this);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 12, Padding = new Padding(22) };
            foreach (int height in new [] { 32, 42, 28, 24, 28, 32, 24, 32, 48, 38, 40 }) panel.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var titleRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = Padding.Empty };
            titleRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); titleRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 36));
            titleRow.Controls.Add(new Label { Text = Text, Font = new Font("Segoe UI Semibold", 15), Dock = DockStyle.Fill }, 0, 0);
            ThemeChoice = Palette.CreateToggle(appearancePath, message => modelHelp.Text = message); titleRow.Controls.Add(ThemeChoice, 1, 0);
            panel.Controls.Add(titleRow, 0, 0);
            panel.Controls.Add(new Label { Text = "Utilisez un compte déjà connecté sur cet ordinateur.\nLa connexion doit être active dans Codex ou Claude Code.", ForeColor = Palette.Muted, Dock = DockStyle.Fill }, 0, 1);
            Chatgpt.Text = "ChatGPT"; Claude.Text = "Claude";
            Chatgpt.Dock = Claude.Dock = DockStyle.Fill; Chatgpt.Enabled = Claude.Enabled = false;
            chatStatus.Text = claudeStatus.Text = "Vérification de la connexion…";
            foreach (var label in new [] { chatStatus, claudeStatus, modelHelp }) { label.Dock = DockStyle.Fill; label.ForeColor = Palette.Muted; label.Font = new Font("Segoe UI", 9); }
            panel.Controls.Add(Chatgpt, 0, 2); panel.Controls.Add(chatStatus, 0, 3); panel.Controls.Add(Claude, 0, 4); panel.Controls.Add(claudeStatus, 0, 5);
            panel.Controls.Add(new Label { Text = "Modèle", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 6);
            Models.Dock = DockStyle.Fill; Models.DropDownStyle = ComboBoxStyle.DropDownList; Models.BackColor = Palette.Surface; Models.ForeColor = Palette.Text; Models.FlatStyle = FlatStyle.Flat;
            Models.DrawMode = DrawMode.OwnerDrawFixed; Models.ItemHeight = 22;
            Models.DrawItem += (s,e) => {
                using (var brush = new SolidBrush(Palette.Surface)) e.Graphics.FillRectangle(brush, e.Bounds);
                if (e.Index >= 0) TextRenderer.DrawText(e.Graphics, Convert.ToString(Models.Items[e.Index]), Models.Font, e.Bounds, Palette.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                if ((e.State & DrawItemState.Focus) != 0) e.DrawFocusRectangle();
            };
            panel.Controls.Add(Models, 0, 7); panel.Controls.Add(modelHelp, 0, 8);
            panel.Controls.Add(new Label { Text = LocalAi.TokenNotice, Dock = DockStyle.Fill, ForeColor = Palette.Text }, 0, 9);
            var buttons = new TableLayoutPanel { ColumnCount = 3, Dock = DockStyle.Fill, Margin = Padding.Empty };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34)); buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33)); buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            var cancel = new SoftButton { Text = current == null ? "Quitter" : "Annuler", DialogResult = DialogResult.Cancel, Dock = DockStyle.Fill };
            RefreshButton.Text = "Actualiser"; RefreshButton.Dock = Accept.Dock = DockStyle.Fill; Accept.Text = persist ? "Enregistrer" : "Valider"; Accept.Enabled = false;
            buttons.Controls.Add(RefreshButton); buttons.Controls.Add(cancel); buttons.Controls.Add(Accept); panel.Controls.Add(buttons, 0, 10); Controls.Add(panel);
            CancelButton = cancel; AcceptButton = Accept;
            Chatgpt.CheckedChanged += (s,e) => { if (Chatgpt.Checked) PopulateModels("chatgpt"); };
            Claude.CheckedChanged += (s,e) => { if (Claude.Checked) PopulateModels("claude"); };
            Models.SelectedIndexChanged += (s,e) => { Accept.Enabled = !checking && Models.SelectedIndex >= 0 && (Chatgpt.Checked && Chatgpt.Enabled || Claude.Checked && Claude.Enabled); };
            RefreshButton.Click += async (s,e) => await RefreshConnections();
            Accept.Click += (s,e) => {
                try { if (SaveSelection()) DialogResult = DialogResult.OK; }
                catch (IOException) { modelHelp.Text = "Enregistrement impossible. Vérifiez l’accès au dossier de l’application."; }
                catch (UnauthorizedAccessException) { modelHelp.Text = "Accès au dossier de paramètres refusé."; }
            };
            if (autoCheck) Shown += async (s,e) => await RefreshConnections();
            HandleCreated += (s,e) => { Palette.ApplyCaption(this); Models.ItemHeight = (int)(24 * CurrentAutoScaleDimensions.Height / 96f); };
            Palette.Apply(this);
            AutoScaleDimensions = new SizeF(96, 96); ResumeLayout(true);
        }
        private void PopulateModels(string provider) {
            Models.Items.Clear();
            if (provider == "chatgpt") { Models.Items.Add("Luna · économique"); Models.Items.Add("Terra · reformulation"); Models.Items.Add("Sol · alternative"); modelHelp.Text = "Luna : correction courante. Terra : reformulations délicates. Sol : comparaison de qualité. Accès selon votre compte."; }
            else { Models.Items.Add("Haiku · économique"); Models.Items.Add("Sonnet · alternative"); modelHelp.Text = "Haiku : correction courte et rapide. Sonnet : alternative pour les reformulations exigeantes. Accès selon votre compte."; }
            Models.SelectedIndex = initial == null || initial.Provider != provider ? 0 : provider == "chatgpt" ? initial.Model == "gpt-6.1-sol" ? 2 : initial.Model == "gpt-6.1-terra" ? 1 : 0 : initial.Model == "sonnet" ? 1 : 0;
        }
        internal void SetConnections(AiConnection[] values) {
            connections = values;
            bool gpt = false, claude = false;
            foreach (var value in values) {
                if (value.Provider == "chatgpt") { gpt = value.Connected; chatStatus.Text = value.Message; }
                if (value.Provider == "claude") { claude = value.Connected; claudeStatus.Text = value.Message; }
            }
            Chatgpt.Enabled = gpt; Claude.Enabled = claude;
            if (Chatgpt.Checked && !gpt) Chatgpt.Checked = false;
            if (Claude.Checked && !claude) Claude.Checked = false;
            if (!Chatgpt.Checked && !Claude.Checked) {
                if (initial != null && initial.Provider == "claude" && claude) Claude.Checked = true;
                else if (gpt) Chatgpt.Checked = true; else if (claude) Claude.Checked = true;
            }
            Models.Enabled = gpt && Chatgpt.Checked || claude && Claude.Checked;
            Accept.Enabled = Models.Enabled && Models.SelectedIndex >= 0 && !checking;
            if (!gpt && !claude) { Models.Items.Clear(); modelHelp.Text = "Aucun compte utilisable. Connectez votre IA dans son application locale, puis Actualiser."; }
        }
        internal bool SaveSelection() {
            if (checking || !Accept.Enabled) return false;
            string provider = Chatgpt.Checked && Chatgpt.Enabled ? "chatgpt" : Claude.Checked && Claude.Enabled ? "claude" : null;
            if (provider == null || Models.SelectedIndex < 0) return false;
            var choice = AiSettings.Default(provider);
            if (provider == "chatgpt") choice.Model = Models.SelectedIndex == 2 ? "gpt-6.1-sol" : Models.SelectedIndex == 1 ? "gpt-6.1-terra" : "gpt-6-luna";
            else if (Models.SelectedIndex == 1) choice.Model = "sonnet";
            if (persist) AiSettingsStore.Save(choice, settingsPath);
            Selected = choice; return true;
        }
        internal async Task RefreshConnections() {
            if (checking) return;
            checking = true; Accept.Enabled = RefreshButton.Enabled = Chatgpt.Enabled = Claude.Enabled = Models.Enabled = false;
            chatStatus.Text = claudeStatus.Text = "Vérification de la connexion…";
            try { var values = await discover(); if (!IsDisposed) { checking = false; SetConnections(values); } }
            catch (Exception) { if (!IsDisposed) { chatStatus.Text = claudeStatus.Text = "Connexion non vérifiable · réessayez."; } }
            finally { checking = false; if (!IsDisposed) RefreshButton.Enabled = true; }
        }
    }
}

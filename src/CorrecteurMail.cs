using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("Korecto")]
[assembly: System.Reflection.AssemblyProduct("Korecto")]
[assembly: System.Reflection.AssemblyVersion(CorrecteurMail.AppBrand.Version + ".0")]
[assembly: System.Reflection.AssemblyFileVersion(CorrecteurMail.AppBrand.Version + ".0")]
[assembly: System.Reflection.AssemblyInformationalVersion(CorrecteurMail.AppBrand.Version)]

namespace CorrecteurMail {
    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleWindow {
        void GetWindow(out IntPtr handle);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
    }
    internal static class Native {
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        private static readonly OutlookProcessCache foregroundProcess = new OutlookProcessCache();
        internal static bool OutlookForeground() {
            uint id; GetWindowThreadProcessId(GetForegroundWindow(), out id);
            return foregroundProcess.IsOutlook(id);
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder name, int size);
        // Outlook's main and message windows are rctrl_renwnd32; a ribbon menu, drop-down or
        // dialog of Outlook in front of the draft is only a passing window.
        internal static bool OutlookPopupForeground() {
            IntPtr window = GetForegroundWindow(); var name = new StringBuilder(64);
            return OutlookForeground() && GetClassName(window, name, name.Capacity) > 0 && name.ToString() != "rctrl_renwnd32";
        }
        internal static void Release(object o) { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
        internal static string Identity(object o) {
            IntPtr handle;
            ((IOleWindow)o).GetWindow(out handle);
            if (handle == IntPtr.Zero) throw new COMException("Fenêtre OUTLOOK inaccessible.");
            uint process; GetWindowThreadProcessId(handle, out process);
            return process.ToString() + ":" + handle.ToInt64().ToString();
        }
    }
    // The 250 ms poll asks for the foreground process several times per tick; it rarely
    // changes, so its name is looked up only when the process id does.
    internal sealed class OutlookProcessCache {
        private readonly Func<uint, string> nameOf;
        private uint lastId;
        private bool known, lastResult;
        internal int Lookups;
        internal OutlookProcessCache(Func<uint, string> processName = null) { nameOf = processName ?? ProcessName; }
        internal bool IsOutlook(uint id) {
            if (known && id == lastId) return lastResult;
            Lookups++;
            string name; try { name = nameOf(id); } catch { name = null; }
            lastId = id; known = true; lastResult = "OUTLOOK".Equals(name, StringComparison.OrdinalIgnoreCase);
            return lastResult;
        }
        private static string ProcessName(uint id) { using (var p = Process.GetProcessById((int)id)) return p.ProcessName; }
    }
    // A changed draft gets a new version, which invalidates outstanding results.
    internal sealed class Debounce {
        internal string Key = "", Text = "";
        internal long Version;
        internal DateTime Changed;
        internal bool Observe(string key, string text, DateTime now) {
            if (key == Key && text == Text) return false;
            Key = key; Text = text; Changed = now; Version++; return true;
        }
        internal bool Matches(long v, string key, string text) { return Version == v && Key == key && Text == text; }
    }
    internal static class Integrity {
        internal static bool SameNumbers(string source, string corrected) {
            var a = Regex.Matches(source, @"(?:[+\-−]\s*)?\d+(?:[.,]\d+)*");
            var b = Regex.Matches(corrected, @"(?:[+\-−]\s*)?\d+(?:[.,]\d+)*");
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++) if (a[i].Value != b[i].Value) return false;
            return true;
        }
    }
    internal sealed class Draft : IDisposable {
        internal static bool SyntheticTest;
        private static bool lastInline;
        internal object Inspector, Item, Document, Range;
        internal string Key, Text;
        internal static Draft Read() {
            if (SyntheticTest) return new Draft { Key = "brouillon-fictif", Text = "Bonjour,\rLes document sont prêt.\rMerci\r" };
            object app = null, inspector = null, item = null, doc = null, range = null, bookmarks = null;
            try {
                app = Marshal.GetActiveObject("Outlook.Application");
                inspector = ((dynamic)app).ActiveInspector();
                bool outlookForeground = Native.OutlookForeground();
                if (!outlookForeground && lastInline) { Native.Release(inspector); inspector = null; }
                string identity = inspector == null ? null : Native.Identity(inspector);
                // ActiveInspector can still refer to the last editor while the inbox is active.
                if (outlookForeground) {
                    uint foregroundProcess; IntPtr foreground = Native.GetForegroundWindow();
                    Native.GetWindowThreadProcessId(foreground, out foregroundProcess);
                    if (identity != foregroundProcess.ToString() + ":" + foreground.ToInt64().ToString()) { Native.Release(inspector); inspector = null; }
                }
                bool inline = inspector == null;
                if (inline) {
                    inspector = ((dynamic)app).ActiveExplorer();
                    if (inspector == null) return null;
                    identity = Native.Identity(inspector);
                    item = ((dynamic)inspector).ActiveInlineResponse;
                    if (item == null) return null;
                    // EntryID is empty for unsaved replies. ConversationIndex is set when the reply
                    // is created and survives new COM wrappers; the live COM identity is a fallback.
                    string conversation = null;
                    try { conversation = Convert.ToString(((dynamic)item).ConversationIndex); }
                    catch (COMException) { } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                    IntPtr itemIdentity = Marshal.GetIUnknownForObject(item);
                    try { identity += InlineIdentity(conversation, itemIdentity.ToInt64()); } finally { Marshal.Release(itemIdentity); }
                } else item = ((dynamic)inspector).CurrentItem;
                if ((int)((dynamic)item).Class != 43 || (bool)((dynamic)item).Sent) return null;
                doc = inline ? ((dynamic)inspector).ActiveInlineResponseWordEditor : ((dynamic)inspector).WordEditor;
                if (doc == null) return null;
                range = ((dynamic)doc).Content;
                bookmarks = ((dynamic)doc).Bookmarks;
                // Leave signature and quoted history untouched when Outlook marks them.
                foreach (string name in new [] { "_MailAutoSig", "_MailOriginal" }) {
                    object bookmark = null, markedRange = null;
                    try {
                        if ((bool)((dynamic)bookmarks).Exists(name)) {
                            bookmark = ((dynamic)bookmarks).Item(name);
                            markedRange = ((dynamic)bookmark).Range;
                            int start = (int)((dynamic)markedRange).Start;
                            if (start >= (int)((dynamic)range).Start && start < (int)((dynamic)range).End) ((dynamic)range).End = start;
                        }
                    } finally { Native.Release(markedRange); Native.Release(bookmark); }
                }
                var result = new Draft { Inspector = inspector, Item = item, Document = doc, Range = range,
                    Key = identity, Text = Convert.ToString(((dynamic)range).Text) ?? "" };
                lastInline = inline;
                inspector = item = doc = range = null;
                return result;
            } catch (Exception) {
                // COM maps some HRESULTs to other exceptions (E_INVALIDARG, E_ACCESSDENIED…). This read
                // runs four times a second: any failure means "no draft", never an error dialog.
                return null;
            } finally { Native.Release(bookmarks); Native.Release(range); Native.Release(doc); Native.Release(item); Native.Release(inspector); Native.Release(app); }
        }
        internal static string InlineIdentity(string conversationIndex, long comIdentity) {
            return ":inline:" + (String.IsNullOrEmpty(conversationIndex) ? comIdentity.ToString() : conversationIndex);
        }
        public void Dispose() { Native.Release(Range); Native.Release(Document); Native.Release(Item); Native.Release(Inspector); }
    }
    internal interface ICorrectionRunner : IDisposable {
        string Progress { get; }
        Task<string> Correct(string text, bool fluent, string advice = null);
        Task<string> Advise(string text, string question);
        void Stop();
    }
    internal sealed class AccountRunner : ICorrectionRunner {
        private Process running;
        private bool disposed;
        private int cancellation;
        private readonly object gate = new object();
        private readonly Func<string> locateCodex;
        private readonly string workingDirectory;
        private readonly int timeoutMs;
        private readonly AiSettings settings;
        private readonly Func<AiConnection> inspectConnection;
        internal const int ConnectionTrustMinutes = 10;
        private DateTime verifiedUntil = DateTime.MinValue;
        internal AccountRunner(Func<string> executableLocator = null, string directory = null, int timeoutMilliseconds = 60000, AiSettings aiSettings = null, Func<AiConnection> connectionInspector = null) {
            settings = (aiSettings ?? AiSettings.Default()).Copy();
            if (!settings.Valid) throw new ArgumentException("Choix d’IA ou de modèle invalide.");
            inspectConnection = connectionInspector;
            locateCodex = executableLocator;
            workingDirectory = directory ?? AppDomain.CurrentDomain.BaseDirectory;
            timeoutMs = timeoutMilliseconds;
            Progress = "Connexion à " + settings.ProviderName + "…";
        }
        public string Progress { get; private set; }
        internal static string ReadCorrection(string line, JavaScriptSerializer json) {
            if (String.IsNullOrWhiteSpace(line)) return null;
            try {
                var e = json.DeserializeObject(line) as Dictionary<string, object>;
                if (e == null || !e.ContainsKey("type") || e["type"] as string != "item.completed" || !e.ContainsKey("item")) return null;
                var item = e["item"] as Dictionary<string, object>;
                if (item == null || !item.ContainsKey("type") || item["type"] as string != "agent_message" || !item.ContainsKey("text") || !(item["text"] is string)) return null;
                var result = json.DeserializeObject((string)item["text"]) as Dictionary<string, object>;
                if (result == null || !result.ContainsKey("texte") || !(result["texte"] is string) || String.IsNullOrWhiteSpace((string)result["texte"])) return null;
                return (string)result["texte"];
            } catch (ArgumentException) { return null; } catch (InvalidOperationException) { return null; }
        }
        internal static string FindCodex() {
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenAI", "Codex", "bin");
            if (Directory.Exists(root)) {
                var dirs = Directory.GetDirectories(root);
                Array.Sort(dirs, (a,b) => Directory.GetLastWriteTimeUtc(b).CompareTo(Directory.GetLastWriteTimeUtc(a)));
                foreach (var dir in dirs) { string exe = Path.Combine(dir, "codex.exe"); if (File.Exists(exe)) return exe; }
            }
            foreach (string candidate in new [] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OpenAI", "Codex", "bin", "codex.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "codex.exe")
            }) if (File.Exists(candidate)) return candidate;
            foreach (string entry in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';')) {
                try { string exe = Path.Combine(entry.Trim('"'), "codex.exe"); if (File.Exists(exe)) return exe; } catch { }
            }
            return null;
        }
        internal static string ReadClaudeCorrection(string line, JavaScriptSerializer json) {
            if (String.IsNullOrWhiteSpace(line)) return null;
            try {
                var e = json.DeserializeObject(line) as Dictionary<string, object>;
                if (e == null || !e.ContainsKey("type") || e["type"] as string != "result") return null;
                if (!e.ContainsKey("is_error") || !(e["is_error"] is bool) || (bool)e["is_error"]) return null;
                if (e.ContainsKey("subtype") && e["subtype"] as string != "success") return null;
                Dictionary<string, object> result = null;
                if (e.ContainsKey("structured_output")) result = e["structured_output"] as Dictionary<string, object>;
                else if (e.ContainsKey("result") && e["result"] is string) result = json.DeserializeObject((string)e["result"]) as Dictionary<string, object>;
                if (result == null || !result.ContainsKey("texte") || !(result["texte"] is string) || String.IsNullOrWhiteSpace((string)result["texte"])) return null;
                return (string)result["texte"];
            } catch (ArgumentException) { return null; } catch (InvalidOperationException) { return null; }
        }
        internal static string ArgumentsFor(AiSettings settings, string directory) {
            if (!settings.Valid) throw new ArgumentException("Modèle invalide.");
            if (settings.Provider == "chatgpt") return "exec --ignore-user-config --ephemeral --skip-git-repo-check --sandbox read-only --disable shell_tool --disable multi_agent --color never --json --model " + settings.Model + " -c " + LocalCommand.Quote("model_reasoning_effort=\"low\"") + " --output-schema " + LocalCommand.Quote(AppAssets.SchemaPath(directory)) + " -";
            string schema = AppAssets.Read(directory, "response.schema.json");
            return "-p --safe-mode --tools \"\" --disallowedTools \"mcp__*\" --strict-mcp-config --mcp-config " + LocalCommand.Quote("{\"mcpServers\":{}}") + " --settings " + LocalCommand.Quote("{\"forceLoginMethod\":\"claudeai\"}") + " --no-chrome --no-session-persistence --max-turns 3 --output-format stream-json --verbose --model " + settings.Model + " --json-schema " + LocalCommand.Quote(schema.Trim());
        }
        private string Failure(string detail, string previous) {
            detail = detail.ToLowerInvariant();
            if (detail.Contains("unknown option") || detail.Contains("unexpected argument") || detail.Contains("unrecognized option")) return "Version du client " + (settings.Provider == "claude" ? "Claude Code" : "Codex") + " incompatible. Mettez votre application locale à jour.";
            if (detail.Contains("rate_limit") || detail.Contains("usage limit") || detail.Contains("rate limit")) return "Limite d’utilisation " + settings.ProviderName + " atteinte. Réessayez après le rétablissement du quota.";
            if (detail.Contains("unauthorized") || detail.Contains("authentication") || detail.Contains("not logged in") || detail.Contains("login required")) return "Connexion " + settings.ProviderName + " à renouveler dans " + (settings.Provider == "claude" ? "Claude Code." : "Codex.");
            if (detail.Contains("model_not_found") || detail.Contains("model is not supported") || detail.Contains("model not available") || detail.Contains("does not have access to model")) return "Modèle " + settings.ModelName + " indisponible pour ce compte. Choisissez un autre modèle dans les paramètres d’IA.";
            return previous;
        }
        public Task<string> Correct(string text, bool fluent, string advice = null) { return Ask(fluent ? "REFORMULER" : "STRICT", text, null, advice); }
        public Task<string> Advise(string text, string question) { return Ask("AVIS", text, question, null); }
        private Task<string> Ask(string mode, string text, string question, string advice) {
            int ticket;
            lock (gate) { if (disposed) throw new ObjectDisposedException("AccountRunner"); ticket = cancellation; Progress = "Connexion à " + settings.ProviderName + "…"; }
            return Task.Run(() => {
                LocalCommand command = locateCodex == null ? LocalAi.Find(settings.Provider) : null;
                if (locateCodex != null) { string exe = locateCodex(); if (exe != null) command = new LocalCommand { FileName = exe }; }
                if (command == null) throw new InvalidOperationException((settings.Provider == "claude" ? "Claude Code" : "Codex") + " est introuvable. Connectez votre compte dans son application locale.");
                string root = workingDirectory;
                // A confirmed login is trusted for a while: one client process per correction instead of two.
                bool verify; lock (gate) verify = DateTime.UtcNow >= verifiedUntil;
                if (verify) {
                    var connection = inspectConnection == null ? LocalAi.Check(settings.Provider, command, root) : inspectConnection();
                    if (!connection.Connected) throw new InvalidOperationException(connection.Message);
                    lock (gate) verifiedUntil = DateTime.UtcNow.AddMinutes(ConnectionTrustMinutes);
                }
                var json = new JavaScriptSerializer();
                string rules = AppAssets.Read(root, "CONSIGNES.txt");
                var data = new Dictionary<string, object> { { "mode", mode }, { "brouillon", text } };
                if (question != null) data["question"] = question;
                if (advice != null) data["avis"] = advice;
                string payload = rules + "\n\nDonnées utilisateur :\n" + json.Serialize(data);
                // No body on command line, no API key, no persisted chat, no user MCP servers.
                var info = command.StartInfo(ArgumentsFor(settings, root), root);
                var p = new Process { StartInfo = info };
                var answer = new TaskCompletionSource<string>();
                // What the AI is doing, in the words the user expects: not "correction" for a reformulation or an avis.
                string doing = mode == "AVIS" ? "l’avis" : mode == "REFORMULER" ? "la reformulation" : "la correction";
                string received = mode == "AVIS" ? "Avis reçu." : mode == "REFORMULER" ? "Reformulation reçue." : "Correction reçue.";
                string failure = "Connexion " + settings.ProviderName + " indisponible. Vérifiez l’application locale et vos limites d’utilisation.";
                // Read completed messages while the process is alive. EOF is not a prerequisite for display.
                p.OutputDataReceived += (sender, args) => {
                    if (String.IsNullOrWhiteSpace(args.Data)) return;
                    try {
                        string correction = settings.Provider == "claude" ? ReadClaudeCorrection(args.Data, json) : ReadCorrection(args.Data, json);
                        if (correction != null) { Progress = received; answer.TrySetResult(correction); }
                        else if (args.Data.Contains("turn.started") || args.Data.Contains("\"type\":\"system\"")) Progress = settings.ProviderName + " prépare " + doing + "…";
                        if (args.Data.Contains("\"error\"") || args.Data.Contains("turn.failed") || args.Data.Contains("\"is_error\":true") || args.Data.Contains("\"is_error\": true")) failure = Failure(args.Data, failure);
                    } catch (ArgumentException) { } catch (InvalidOperationException) { }
                };
                p.ErrorDataReceived += (sender, args) => {
                    if (String.IsNullOrWhiteSpace(args.Data)) return;
                    failure = Failure(args.Data, failure);
                };
                lock (gate) {
                    if (disposed || ticket != cancellation) { p.Dispose(); throw new InvalidOperationException("Correction annulée."); }
                    if (running != null) { p.Dispose(); throw new InvalidOperationException("Une demande est déjà en cours."); }
                    // Start under the same lock as Stop: closing the window cannot miss this process.
                    try { p.Start(); running = p; } catch { p.Dispose(); throw; }
                }
                try {
                    var elapsed = Stopwatch.StartNew();
                    p.BeginOutputReadLine(); p.BeginErrorReadLine();
                    byte[] bytes = new UTF8Encoding(false).GetBytes(payload);
                    try { p.StandardInput.BaseStream.Write(bytes, 0, bytes.Length); p.StandardInput.Close(); }
                    catch (IOException) {
                        // The client quit before reading (rejected option, expired login): report its own error.
                        if (p.WaitForExit(2000)) p.WaitForExit();
                        throw new InvalidOperationException(failure);
                    }
                    while (!answer.Task.Wait(150)) {
                        lock (gate) { if (disposed || ticket != cancellation) throw new InvalidOperationException("Correction annulée."); }
                        if (elapsed.Elapsed.TotalMilliseconds >= timeoutMs) throw new InvalidOperationException("Aucune réponse reçue après " + (timeoutMs / 1000.0).ToString("0.#") + " secondes. Cliquez sur le bouton pour réessayer.");
                        if (p.HasExited) {
                            // Allow the async output callback to consume the last buffered line.
                            if (answer.Task.Wait(1200)) break;
                            throw new InvalidOperationException(p.ExitCode != 0 ? failure : settings.ProviderName + " n’a pas renvoyé de réponse exploitable. Cliquez à nouveau sur le bouton.");
                        }
                    }
                    lock (gate) { if (disposed || ticket != cancellation) throw new InvalidOperationException("Correction annulée."); }
                    return answer.Task.Result;
                } catch {
                    // Any real failure may mean an expired login: check it again before the next correction.
                    lock (gate) { if (!disposed && ticket == cancellation) verifiedUntil = DateTime.MinValue; }
                    throw;
                } finally {
                    lock (gate) {
                        try { if (!p.HasExited) p.Kill(); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
                        if (running == p) running = null;
                    }
                    p.Dispose();
                }
            });
        }
        public void Stop() {
            lock (gate) {
                cancellation++;
                if (running == null) return;
                try { if (!running.HasExited) running.Kill(); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
            }
        }
        public void Dispose() { lock (gate) { disposed = true; Stop(); } }
    }
    internal static class Palette {
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        internal static void ApplyCaption(Form form) {
            if (!form.IsHandleCreated) return;
            // Windows 11 supports caption colours; older Windows ignores these attributes.
            int dark = IsDark ? 1 : 0, background = ColorTranslator.ToWin32(Background), foreground = ColorTranslator.ToWin32(Text);
            DwmSetWindowAttribute(form.Handle, 20, ref dark, 4);
            DwmSetWindowAttribute(form.Handle, 35, ref background, 4);
            DwmSetWindowAttribute(form.Handle, 36, ref foreground, 4);
        }
        internal static bool IsDark { get; private set; }
        static Palette() { IsDark = true; }
        internal static Color Background { get { return IsDark ? Color.FromArgb(22, 29, 29) : Color.FromArgb(245, 247, 244); } }
        internal static Color Surface { get { return IsDark ? Color.FromArgb(32, 41, 40) : Color.White; } }
        internal static Color Border { get { return IsDark ? Color.FromArgb(56, 72, 68) : Color.FromArgb(217, 227, 218); } }
        internal static Color Text { get { return IsDark ? Color.FromArgb(235, 242, 236) : Color.FromArgb(33, 52, 45); } }
        internal static Color Muted { get { return IsDark ? Color.FromArgb(166, 184, 174) : Color.FromArgb(97, 115, 104); } }
        internal static Color Hover { get { return IsDark ? Color.FromArgb(45, 60, 54) : Color.FromArgb(233, 241, 232); } }
        internal static Color Accent { get { return IsDark ? Color.FromArgb(166, 215, 174) : Color.FromArgb(51, 107, 74); } }
        internal static Color AccentText { get { return IsDark ? Color.FromArgb(22, 45, 30) : Color.White; } }
        internal static Color PrimaryHover { get { return IsDark ? Color.FromArgb(191, 232, 197) : Color.FromArgb(38, 88, 58); } }
        internal static Color Disabled { get { return IsDark ? Color.FromArgb(125, 144, 132) : Color.FromArgb(123, 139, 128); } }
        internal static void Use(bool dark) {
            IsDark = dark;
            foreach (Form form in Application.OpenForms) Apply(form);
        }
        internal static Button CreateToggle(string path, Action<string> showError) {
            var button = new SoftButton { Dock = DockStyle.Fill, Tag = "theme-toggle", Margin = new Padding(3, 1, 0, 1), Font = new Font("Segoe UI", 8.5f) };
            button.Text = "";
            button.AccessibleName = "Passer au mode " + (IsDark ? "clair" : "sombre");
            var hint = new ToolTip { ShowAlways = true };
            button.MouseEnter += (s,e) => hint.SetToolTip(button, button.AccessibleName);
            button.Disposed += (s,e) => hint.Dispose();
            button.Click += (s,e) => {
                try { bool dark = !IsDark; AppearanceStore.Save(dark, path); Use(dark); }
                catch (IOException) { showError("Thème non enregistré · vérifiez l’accès au dossier."); }
                catch (UnauthorizedAccessException) { showError("Thème non enregistré · accès au dossier refusé."); }
            };
            return button;
        }
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] private static extern int SetWindowTheme(IntPtr window, string appName, string idList);
        // Scrollbars of text boxes follow the theme (dark scrollbars need the DarkMode_Explorer visual style).
        internal static void ThemeScrollbars(Control box) {
            if (box.IsHandleCreated) SetWindowTheme(box.Handle, IsDark ? "DarkMode_Explorer" : "Explorer", null);
        }
        internal static void Apply(Control control) {
            bool muted = control.Tag as string == "theme-muted" || control.ForeColor.ToArgb() == Muted.ToArgb();
            if (control is Label && muted) control.Tag = "theme-muted";
            control.BackColor = control is TextBox || control is RichTextBox || control is TabPage || control is ComboBox || control.Tag as string == "surface" ? Surface : Background;
            control.ForeColor = control is Label && muted ? Muted : Text;
            if (control.Tag as string == "theme-toggle") {
                control.Text = "";
                control.AccessibleName = "Passer au mode " + (IsDark ? "clair" : "sombre");
                control.AccessibleDescription = "Thème " + (IsDark ? "sombre" : "clair") + " actif.";
            }
            if (control is TextBoxBase) ThemeScrollbars(control);
            foreach (Control child in control.Controls) Apply(child);
            var form = control as Form; if (form != null) ApplyCaption(form);
            control.Invalidate(true);
        }
    }
    // A text box that shows a muted hint while it is empty (the native hint does not exist for multiline boxes).
    internal sealed class HintTextBox : TextBox {
        internal string Hint = "";
        internal Font HintFont;
        protected override void WndProc(ref Message m) {
            base.WndProc(ref m);
            if (m.Msg != 0x000F || TextLength > 0 || Hint.Length == 0) return;
            using (var graphics = Graphics.FromHwnd(Handle))
                TextRenderer.DrawText(graphics, Hint, HintFont ?? Font, new Rectangle(1, 1, Math.Max(0, ClientSize.Width - 2), Math.Max(0, ClientSize.Height - 2)), Palette.Muted, TextFormatFlags.Left | TextFormatFlags.Top | TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
        }
    }
    internal sealed class SoftButton : Button {
        internal bool Primary;
        private bool hover;
        internal SoftButton() {
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; Cursor = Cursors.Hand;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e) {
            e.Graphics.Clear(Parent == null ? Palette.Background : Parent.BackColor);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new Rectangle(1, 1, Width - 3, Height - 3);
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            using (var path = new GraphicsPath()) {
                const int diameter = 14;
                path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
                path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
                path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90); path.CloseFigure();
                Color fill = Primary && Enabled ? (hover ? Palette.PrimaryHover : Palette.Accent) : (hover && Enabled ? Palette.Hover : Palette.Surface);
                using (var brush = new SolidBrush(fill)) e.Graphics.FillPath(brush, path);
                using (var pen = new Pen(Palette.Border)) e.Graphics.DrawPath(pen, path);
            }
            if (Tag as string == "theme-toggle" || Tag as string == "icon-settings" || Tag as string == "icon-pause") DrawSymbol(e.Graphics, bounds);
            else TextRenderer.DrawText(e.Graphics, Text, Font, bounds, !Enabled ? Palette.Disabled : Primary ? Palette.AccentText : Palette.Text, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            if (Focused && ShowFocusCues) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -4, -4), Palette.Muted, Palette.Surface);
        }
        private void DrawSymbol(Graphics graphics, Rectangle bounds) {
            float dpi = graphics.DpiX / 96f;
            float size = Math.Min(17f * dpi, Math.Min(bounds.Width, bounds.Height) - 8f * dpi);
            if (size <= 0) return;
            var saved = graphics.Save();
            graphics.TranslateTransform(bounds.Left + bounds.Width / 2f, bounds.Top + bounds.Height / 2f);
            graphics.ScaleTransform(size / 24f, size / 24f);
            Color color = Enabled ? Palette.Muted : Palette.Disabled;
            using (var pen = new Pen(color, 1.8f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            using (var brush = new SolidBrush(color)) {
                string symbol = Tag as string;
                if (symbol == "icon-pause") {
                    if (Text == "▶") graphics.FillPolygon(brush, new [] { new PointF(-4.5f, -8), new PointF(8, 0), new PointF(-4.5f, 8) });
                    else { graphics.DrawLine(pen, -4, -7, -4, 7); graphics.DrawLine(pen, 4, -7, 4, 7); }
                } else if (symbol == "icon-settings") {
                    using (var gear = new GraphicsPath()) {
                        var points = new PointF[32];
                        for (int i = 0; i < points.Length; i++) {
                            double angle = (i - .5) * Math.PI / 16;
                            float radius = i % 4 == 1 || i % 4 == 2 ? 10 : 8;
                            points[i] = new PointF((float)Math.Cos(angle) * radius, (float)Math.Sin(angle) * radius);
                        }
                        gear.AddPolygon(points); graphics.DrawPath(pen, gear); graphics.DrawEllipse(pen, -3, -3, 6, 6);
                    }
                } else if (Palette.IsDark) {
                    graphics.DrawEllipse(pen, -4, -4, 8, 8);
                    for (int i = 0; i < 8; i++) {
                        double angle = Math.PI * i / 4;
                        graphics.DrawLine(pen, (float)Math.Cos(angle) * 7, (float)Math.Sin(angle) * 7, (float)Math.Cos(angle) * 10, (float)Math.Sin(angle) * 10);
                    }
                } else {
                    // One antialiased contour: no jagged region subtraction or font glyph.
                    using (var moon = new GraphicsPath()) {
                        moon.AddBezier(3.5f, -9, -9, -11, -14, 5, -3, 9);
                        moon.AddBezier(-3, 9, 1, 11, 6, 9, 8.5f, 5.5f);
                        moon.AddBezier(8.5f, 5.5f, -.5f, 8, -4, -3, 3.5f, -9);
                        moon.CloseFigure(); graphics.FillPath(brush, moon);
                    }
                }
            }
            graphics.Restore(saved);
        }
    }
    internal sealed class DarkTabs : TabControl {
        internal DarkTabs() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            ItemSize = new Size(110, 30); SizeMode = TabSizeMode.Fixed;
        }
        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            using (var graphics = CreateGraphics()) ItemSize = new Size((int)(112 * graphics.DpiX / 96), (int)(34 * graphics.DpiY / 96));
        }
        protected override void OnPaint(PaintEventArgs e) {
            e.Graphics.Clear(Palette.Background);
            for (int i = 0; i < TabPages.Count; i++) {
                var rect = GetTabRect(i);
                TextRenderer.DrawText(e.Graphics, TabPages[i].Text, Font, rect, i == SelectedIndex ? Palette.Text : Available(TabPages[i]) ? Palette.Muted : Palette.Border, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                if (i == SelectedIndex) using (var pen = new Pen(Palette.Accent, 3)) e.Graphics.DrawLine(pen, rect.Left + 10, rect.Bottom - 2, rect.Right - 10, rect.Bottom - 2);
            }
        }
        protected override void OnSelectedIndexChanged(EventArgs e) { Invalidate(); base.OnSelectedIndexChanged(e); }
        internal Func<TabPage, bool> Available = page => true;
        protected override void OnSelecting(TabControlCancelEventArgs e) { if (e.TabPage != null && !Available(e.TabPage)) e.Cancel = true; base.OnSelecting(e); }
    }
    internal sealed class Popup : Form {
        private readonly Label status = new Label();
        private readonly Label context = new Label();
        private readonly Label info = new Label();
        private readonly Button ai = new SoftButton();
        private readonly Button updateAvailable = new SoftButton();
        private readonly TableLayoutPanel noticeRow = new TableLayoutPanel();
        private readonly Button theme;
        private readonly TextBox result = new TextBox(), reformulated = new TextBox();
        private readonly HintTextBox adviceQuestion = new HintTextBox(), adviceBox = new HintTextBox();
        private readonly Button adviceAsk = new SoftButton(), adviceReformulate = new SoftButton();
        private readonly List<Button> toneButtons = new List<Button>();
        private readonly RichTextBox resultView = new RichTextBox();
        private readonly TextBox manualInput = new TextBox();
        private readonly TabControl textTabs = new DarkTabs();
        private readonly TabPage inputPage = new TabPage("Texte d’origine"), resultPage = new TabPage("Corrigé"), reformulationPage = new TabPage("Reformulé"), advicePage = new TabPage("Avis");
        private readonly Button manual = new SoftButton(), copy = new SoftButton();
        private readonly Button expand = new SoftButton();
        private readonly TableLayoutPanel layout = new TableLayoutPanel();
        private readonly ToolTip tips = new ToolTip();
        private readonly Button apply = new SoftButton(), reformulate = new SoftButton(), retry = new SoftButton { Primary = true }, pause = new SoftButton();
        private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer(), updateTimer = new System.Windows.Forms.Timer();
        private ReleaseInfo availableUpdate;
        private readonly ModePreference modePref;
        private bool defaultManual, awayArmed;
        private SelectionAccess selectionAccess;
        private readonly bool internalMode;
        private readonly Debounce state = new Debounce();
        private readonly Dictionary<string, int> draftNumbers = new Dictionary<string, int>();
        // Texts Korecto already handled, per draft: a source and its strict correction, or a text
        // it applied (null). Ctrl+Z / Ctrl+Y in Outlook bring them back without a new request.
        private readonly Dictionary<string, string> handled = new Dictionary<string, string>();
        private ICorrectionRunner runner;
        private AiSettings settings;
        private readonly Func<Draft> readDraft;
        private readonly Func<bool> outlookForeground, outlookPopup;
        private readonly NotifyIcon tray = new NotifyIcon();
        private readonly Icon notificationIcon;
        private readonly bool testMode;
        private bool hiddenByUser, exitRequested, resourcesReleased;
        private bool suppressInitialVisibility;
        private bool paused, busy, closing, detected, expanded, manualMode, settingsOpen;
        private long requested = -1, resultVersion = -1, correctedVersion = -1;
        private long correctionVersion = -1, reformulationVersion = -1, adviceVersion = -1;
        private string manualDraft = "";
        private bool updatingInput;
        private Size expandedClientSize = new Size(520, 350);
        private bool CanCorrect { get { return !busy && !paused && !String.IsNullOrWhiteSpace(state.Text) && correctedVersion != state.Version; } }
        private string resultKey, resultSource;
        // The tabs hold a real proposal (not an error) for this draft, possibly from an earlier text.
        private bool resultKept, reformulationKept;
        private string keptKey;
        private DateTime requestStarted;
        private string requestTitle;
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); Palette.ApplyCaption(this); }
        protected override void OnFormClosing(FormClosingEventArgs e) {
            if (!exitRequested && e.CloseReason == CloseReason.UserClosing) {
                e.Cancel = true; hiddenByUser = true; Hide(); ShowInTaskbar = false;
            } else closing = true;
            base.OnFormClosing(e);
        }
        internal void RestoreFromTray() {
            if (closing || IsDisposed) return;
            hiddenByUser = false; ShowInTaskbar = true; WindowState = FormWindowState.Normal; Show();
            if (!testMode) { timer.Start(); Activate(); }
        }
        internal void ExitApplication() {
            if (IsDisposed || exitRequested) return;
            exitRequested = true; closing = true; Close();
        }
        private void ReleaseResources() {
            if (resourcesReleased) return;
            resourcesReleased = true; closing = true; timer.Stop();
            if (selectionAccess != null) selectionAccess.Dispose();
            if (runner != null) runner.Dispose();
            tray.Visible = false; tray.Dispose(); timer.Dispose(); updateTimer.Dispose(); tips.Dispose();
        }
        protected override void Dispose(bool disposing) {
            if (disposing) ReleaseResources();
            base.Dispose(disposing);
        }
        protected override void SetVisibleCore(bool value) {
            if (value && suppressInitialVisibility) { suppressInitialVisibility = false; if (!IsHandleCreated) CreateHandle(); base.SetVisibleCore(false); return; }
            base.SetVisibleCore(value);
        }
        internal Popup(bool testMode = false, ICorrectionRunner correctionRunner = null, Func<Draft> draftReader = null, Func<bool> foregroundReader = null, AiSettings aiSettings = null, string appearancePath = null, bool startInTray = false, bool internalMode = false, Func<bool> popupReader = null, ModePreference modePreference = null) {
            SuspendLayout();
            this.testMode = testMode;
            this.internalMode = internalMode;
            modePref = modePreference ?? new ModePreference(); defaultManual = modePref.ManualByDefault;
            suppressInitialVisibility = startInTray;
            settings = (aiSettings ?? AiSettings.Default()).Copy();
            runner = correctionRunner ?? new AccountRunner(aiSettings: settings);
            readDraft = draftReader ?? Draft.Read;
            outlookForeground = foregroundReader ?? Native.OutlookForeground;
            outlookPopup = popupReader ?? Native.OutlookPopupForeground;
            Text = AppBrand.Name;
            AppBrand.SetIcon(this);
            Font = new Font("Segoe UI", 9.5f);
            DoubleBuffered = true;
            AutoScaleDimensions = new SizeF(96, 96);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(520, 160); MinimumSize = SizeFromClientSize(new Size(480, 160));
            StartPosition = FormStartPosition.Manual;
            var area = Screen.PrimaryScreen.WorkingArea; Location = new Point(area.Right - Width - 20, area.Bottom - Height - 20);
            TopMost = true; BackColor = Palette.Background; ForeColor = Palette.Text;
            layout.Dock = DockStyle.Fill; layout.ColumnCount = 1; layout.RowCount = 5; layout.Padding = new Padding(12, 6, 12, 6);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            var brand = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, Margin = Padding.Empty };
            brand.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); brand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); brand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34)); brand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34)); brand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34));
            brand.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            brand.Controls.Add(new Label { Text = AppBrand.Name, Font = new Font("Segoe UI Semibold", 20), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true }, 0, 0);
            context.Text = "OUTLOOK"; context.Dock = DockStyle.Fill; context.TextAlign = ContentAlignment.MiddleRight; context.ForeColor = Palette.Muted; context.Font = new Font("Segoe UI", 8); context.AutoEllipsis = true;
            brand.Controls.Add(context, 1, 0);
            theme = Palette.CreateToggle(appearancePath, message => status.Text = message); theme.Margin = new Padding(3, 6, 0, 6); brand.Controls.Add(theme, 2, 0);
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            status.Dock = DockStyle.Fill; status.Text = "Ouvrez un brouillon OUTLOOK.\nPuis cliquez sur Corriger.";
            status.AutoEllipsis = true; status.ForeColor = Palette.Muted; status.TextAlign = ContentAlignment.MiddleLeft;
            status.TextChanged += (s,e) => tips.SetToolTip(status, status.Text);
            expand.Text = "Aperçu"; expand.Dock = DockStyle.Fill; expand.Margin = new Padding(6, 9, 0, 12); expand.Click += (s,e) => SetExpanded(!expanded);
            tips.SetToolTip(expand, "Afficher ou réduire la proposition. La fenêtre reste redimensionnable.");
            manual.Text = "Manuel"; manual.Dock = DockStyle.Fill; manual.Margin = new Padding(6, 6, 0, 6); manual.Click += (s,e) => SwitchMode();
            tips.SetToolTip(manual, "Coller un texte manuellement ou revenir à la lecture OUTLOOK.");
            header.Controls.Add(status, 0, 0); header.Controls.Add(manual, 1, 0);
            result.Dock = DockStyle.Fill; result.Multiline = true; result.ReadOnly = true; result.ScrollBars = ScrollBars.Vertical; result.BackColor = Palette.Surface; result.ForeColor = Palette.Text; result.BorderStyle = BorderStyle.None; result.MaxLength = 0;
            result.ShortcutsEnabled = true;
            reformulated.Dock = DockStyle.Fill; reformulated.Multiline = true; reformulated.ReadOnly = true; reformulated.ScrollBars = ScrollBars.None; reformulated.BorderStyle = BorderStyle.None; reformulated.MaxLength = 0;
            result.Font = manualInput.Font = reformulated.Font = new Font("Segoe UI", 11);
            manualInput.Dock = DockStyle.Fill; manualInput.Multiline = true; manualInput.ScrollBars = ScrollBars.None; manualInput.AcceptsReturn = true; manualInput.MaxLength = 0;
            manualInput.BackColor = Palette.Surface; manualInput.ForeColor = Palette.Text; manualInput.BorderStyle = BorderStyle.None;
            manualInput.ReadOnly = true;
            foreach (var page in new [] { inputPage, resultPage, reformulationPage, advicePage }) { page.BackColor = Palette.Surface; page.ForeColor = Palette.Text; page.Padding = new Padding(12); }
            manualInput.TextChanged += (s,e) => {
                if (!manualMode || updatingInput) return;
                manualDraft = manualInput.Text;
                state.Observe("manuel", manualInput.Text, DateTime.UtcNow); InvalidateResult(false, !String.IsNullOrWhiteSpace(manualInput.Text));
                if (busy) runner.Stop();
                retry.Enabled = CanCorrect;
                apply.Enabled = !busy && !paused;
                reformulate.Enabled = !busy && !paused && !String.IsNullOrWhiteSpace(manualInput.Text); textTabs.SelectedTab = inputPage;
                status.Text = "Texte manuel · cliquez sur Corriger.";
            };
            // Corrigé shows the proposal with the changed words highlighted; the hidden text box stays the text that is copied and applied.
            result.Visible = false;
            foreach (var text in new [] { manualInput, reformulated, adviceBox }) { var box = text; box.TextChanged += (s,e) => FitScrollbar(box); box.SizeChanged += (s,e) => FitScrollbar(box); }
            foreach (var box in new Control[] { manualInput, reformulated, resultView, adviceBox }) { var scrolling = box; scrolling.HandleCreated += (s,e) => Palette.ThemeScrollbars(scrolling); }
            resultView.Dock = DockStyle.Fill; resultView.ReadOnly = true; resultView.ScrollBars = RichTextBoxScrollBars.Vertical; resultView.BorderStyle = BorderStyle.None; resultView.DetectUrls = false; resultView.Font = new Font("Segoe UI", 11);
            resultView.BackColorChanged += (s,e) => RefreshChanges();
            result.TextChanged += (s,e) => resultView.Text = result.Text;
            inputPage.Controls.Add(manualInput); resultPage.Controls.Add(resultView); resultPage.Controls.Add(result); reformulationPage.Controls.Add(reformulated);
            // Avis: a question about the draft, the answer, then a reformulation that follows it or a tone shortcut. Nothing leaves before a click.
            // Empty field with a muted hint; an empty question asks for a general opinion.
            adviceQuestion.Hint = "Votre question (vide = avis général)";
            adviceBox.HintFont = new Font("Segoe UI", 9.5f);
            adviceBox.Hint = "Posez une question sur votre mail. Exemples : « Le ton est-il adapté ? », « Le message est-il clair ? », « Est-il trop direct ? ». Vous pourrez ensuite le reformuler dans le sens de la réponse.";
            adviceQuestion.MaxLength = 300; adviceQuestion.Font = new Font("Segoe UI", 10); adviceQuestion.BorderStyle = BorderStyle.FixedSingle;
            adviceQuestion.Dock = DockStyle.Fill; adviceQuestion.Margin = new Padding(0, 4, 8, 0);
            adviceQuestion.KeyDown += (s,e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; adviceAsk.PerformClick(); } };
            adviceAsk.Text = "Demander l’avis"; adviceAsk.Dock = DockStyle.Fill; adviceAsk.Margin = Padding.Empty; adviceAsk.Enabled = false;
            adviceAsk.Click += async (s,e) => await RequestAdvice();
            tips.SetToolTip(adviceAsk, "Envoie le brouillon et votre question à l’IA. Rien n’est modifié dans le mail.");
            adviceBox.Dock = DockStyle.Fill; adviceBox.Multiline = true; adviceBox.ReadOnly = true; adviceBox.ScrollBars = ScrollBars.None; adviceBox.BorderStyle = BorderStyle.None; adviceBox.MaxLength = 0; adviceBox.Font = new Font("Segoe UI", 11); adviceBox.Margin = new Padding(0, 8, 0, 8);
            adviceReformulate.Text = "Reformuler selon cet avis"; adviceReformulate.Dock = DockStyle.Fill; adviceReformulate.Margin = Padding.Empty; adviceReformulate.Enabled = false;
            adviceReformulate.Click += async (s,e) => await RequestCore(true, adviceBox.Text);
            var tones = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Margin = new Padding(0, 6, 0, 0), Tag = "surface" };
            tones.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            foreach (var tone in new [] {
                new [] { "Plus ferme", "Reformule avec un ton plus ferme et plus direct, sans devenir agressif ni impoli." },
                new [] { "Plus doux", "Reformule avec un ton plus doux et plus chaleureux, sans devenir familier." },
                new [] { "Plus court", "Raccourcis le texte : garde l'essentiel, supprime les répétitions et les formules inutiles." },
                new [] { "Plus formel", "Reformule avec un ton plus formel et plus soutenu, sans lourdeur." } }) {
                string instruction = tone[1];
                var toneButton = new SoftButton { Text = tone[0], Dock = DockStyle.Fill, Margin = new Padding(0, 0, 6, 0), Font = new Font("Segoe UI", 9), Enabled = false };
                toneButton.Click += async (s,e) => await RequestCore(true, instruction);
                tones.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); tones.Controls.Add(toneButton); toneButtons.Add(toneButton);
            }
            var adviceLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4, Margin = Padding.Empty, Tag = "surface" };
            adviceLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); adviceLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            adviceLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36)); adviceLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); adviceLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); adviceLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            adviceLayout.Controls.Add(adviceQuestion, 0, 0); adviceLayout.Controls.Add(adviceAsk, 1, 0);
            adviceLayout.Controls.Add(adviceBox, 0, 1); adviceLayout.SetColumnSpan(adviceBox, 2);
            adviceLayout.Controls.Add(adviceReformulate, 0, 2); adviceLayout.SetColumnSpan(adviceReformulate, 2);
            adviceLayout.Controls.Add(tones, 0, 3); adviceLayout.SetColumnSpan(tones, 2);
            advicePage.Controls.Add(adviceLayout);
            textTabs.Dock = DockStyle.Fill; textTabs.TabPages.AddRange(new [] { inputPage, resultPage, reformulationPage, advicePage }); textTabs.Visible = false;
            // Each tab opens only when it has something to show; Avis as soon as there is a text.
            ((DarkTabs)textTabs).Available = page => page == inputPage || (page == advicePage ? !String.IsNullOrWhiteSpace(state.Text) : page == resultPage ? result.TextLength > 0 : reformulated.TextLength > 0);
            result.TextChanged += (s,e) => textTabs.Invalidate(); reformulated.TextChanged += (s,e) => textTabs.Invalidate(); manualInput.TextChanged += (s,e) => textTabs.Invalidate();
            // Opening a tab never calls the AI: only the Reformuler button asks for a reformulation.
            textTabs.SelectedIndexChanged += (s,e) => {
                if (textTabs.SelectedTab == advicePage) EnsureHeight(470);
                UpdateProposalActions();
                if (textTabs.SelectedTab == reformulationPage && reformulationVersion != state.Version && resultVersion == state.Version && !busy && !paused) status.Text = "Cliquez sur Reformuler pour proposer une autre tournure.";
            };
            var actions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Margin = Padding.Empty };
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            apply.Text = "Appliquer"; reformulate.Text = "Reformuler"; retry.Text = "Corriger"; pause.Text = "Ⅱ";
            foreach (var b in new [] { retry, apply, reformulate, copy }) { actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); b.Dock = DockStyle.Fill; b.Margin = new Padding(0, 5, 6, 5); actions.Controls.Add(b); }
            pause.Tag = "icon-pause"; pause.Dock = DockStyle.Fill; pause.Margin = new Padding(3, 6, 0, 6); pause.Font = new Font("Segoe UI Symbol", 11); pause.AccessibleName = "Mettre en pause";
            tips.SetToolTip(pause, "Mettre en pause la lecture d’OUTLOOK et les corrections."); brand.Controls.Add(pause, 3, 0);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            info.Dock = DockStyle.Fill; info.Text = LocalAi.TokenNotice; info.ForeColor = Palette.Muted; info.Font = new Font("Segoe UI", 8); info.TextAlign = ContentAlignment.MiddleLeft; info.AutoEllipsis = true;
            ai.Tag = "icon-settings"; ai.Text = "⚙"; ai.Font = new Font("Segoe UI Symbol", 12); ai.AccessibleName = "Paramètres"; ai.Dock = DockStyle.Fill; ai.Margin = new Padding(3, 6, 0, 6); ai.Click += (s,e) => OpenSettings();
            brand.Controls.Add(ai, 4, 0);
            copy.Text = "Copier"; copy.Dock = DockStyle.Fill; copy.Margin = new Padding(0, 5, 0, 5); copy.Enabled = false; copy.Click += (s,e) => CopyResult(null);
            tips.SetToolTip(copy, "Copier toute la correction pour la coller vous-même dans le mail.");
            var content = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty }; content.Controls.Add(textTabs);
            layout.RowStyles.Clear();
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            noticeRow.Dock = DockStyle.Fill; noticeRow.Margin = Padding.Empty; noticeRow.ColumnCount = 2; noticeRow.RowCount = 1;
            noticeRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); noticeRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 0));
            noticeRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); noticeRow.Controls.Add(info, 0, 0);
            updateAvailable.Text = "Mise à jour disponible"; updateAvailable.AccessibleName = updateAvailable.Text;
            updateAvailable.Font = new Font("Segoe UI", 8); updateAvailable.Dock = DockStyle.Fill; updateAvailable.Margin = Padding.Empty;
            updateAvailable.Visible = false;
            updateAvailable.Click += async (s,e) => {
                var pending = availableUpdate;
                if (pending == null || closing) return;
                if (internalMode) { status.Text = "Aperçu du bouton : aucune mise à jour téléchargée en mode test."; return; }
                updateAvailable.Enabled = false;
                try { await UpdateUi.Offer(this, pending, ExitApplication); }
                finally { if (!IsDisposed && !closing) updateAvailable.Enabled = true; }
            };
            noticeRow.Controls.Add(updateAvailable, 1, 0);
            layout.Controls.Add(noticeRow, 0, 0); layout.Controls.Add(brand, 0, 1); layout.Controls.Add(header, 0, 2); layout.Controls.Add(content, 0, 3); layout.Controls.Add(actions, 0, 4); Controls.Add(layout);
            apply.Enabled = reformulate.Enabled = retry.Enabled = false;
            expand.Enabled = false;
            apply.Click += (s,e) => { if (manualMode) PasteManual(); else Apply(); }; reformulate.Click += async (s,e) => await Request(true); retry.Click += async (s,e) => await Request(false);
            pause.Click += (s,e) => TogglePause();
            var menu = new ContextMenuStrip(); menu.Items.Add("Afficher", null, (s,e) => RestoreFromTray()); menu.Items.Add("Pause / Reprendre", null, (s,e) => TogglePause()); menu.Items.Add("Quitter", null, (s,e) => ExitApplication());
            menu.Items.Insert(1, new ToolStripMenuItem("Mode manuel", null, (s,e) => { if (!manualMode) SwitchMode(); RestoreFromTray(); }));
            menu.Items.Insert(2, new ToolStripMenuItem("Paramètres…", null, (s,e) => OpenSettings()));
            notificationIcon = AppBrand.LoadNotificationIcon(16);
            tray.Icon = notificationIcon; tray.Text = AppBrand.Name; tray.ContextMenuStrip = menu; tray.Visible = !testMode; tray.DoubleClick += (s,e) => RestoreFromTray();
            Disposed += (s,e) => notificationIcon.Dispose();
            UpdateAiLabel();
            Palette.Apply(this);
            timer.Interval = 250; timer.Tick += async (s,e) => await Tick(); if (!testMode) timer.Start();
            if (!testMode && !internalMode) {
                updateTimer.Interval = 8000;
                updateTimer.Tick += async (s,e) => { updateTimer.Stop(); await BackgroundUpdateCheck(); };
                if (UpdateUi.IsInstalledCopy() && new UpdatePreference().Enabled) { Updater.CleanOldDownloads(); updateTimer.Start(); }
            }
            FormClosed += (s,e) => ReleaseResources();
            AutoScaleDimensions = new SizeF(96, 96); ResumeLayout(true);
            Location = new Point(area.Right - Width - 20, area.Bottom - Height - 20);
            if (!testMode) selectionAccess = new SelectionAccess(ImportSelection, message => { if (!closing) status.Text = message; });
            if (defaultManual) SwitchMode();
            if (internalMode) {
                SetAvailableUpdate(new ReleaseInfo { Version = new Version(99, 0, 0) });
                Shown += (s,e) => { if (!manualMode) ImportSelection("Bonjour,\r\n\r\nCe texte fictif permet de découvrir la version interne de Korecto.\r\n\r\nVous pouvez aussi sélectionner un texte dans une autre application, puis appuyer sur Alt + K."); };
            }
        }
        // Silent startup check: only an inline button, never a balloon or dialog.
        private async Task BackgroundUpdateCheck(Func<string, string> fetch = null) {
            try {
                var release = await Updater.CheckAsync(fetch);
                if (closing || IsDisposed) return;
                SetAvailableUpdate(release);
            } catch (Exception ex) { if (!Updater.IsKnownError(ex)) throw; }
        }
        internal void SetAvailableUpdate(ReleaseInfo release) {
            availableUpdate = release != null && Updater.IsNewer(AppBrand.Version, release.Version) ? release : null;
            noticeRow.ColumnStyles[1].Width = availableUpdate == null ? 0 : 154 * CurrentAutoScaleDimensions.Width / 96f;
            updateAvailable.Visible = availableUpdate != null;
            tips.SetToolTip(updateAvailable, availableUpdate == null ? "" : "Installer Korecto " + availableUpdate.Version.ToString(3));
        }
        private void UpdateAiLabel() {
            Text = AppBrand.Name + " · " + settings.ProviderName + " · " + settings.ModelName;
            if (internalMode) Text += " · TEST INTERNE";
            tips.SetToolTip(ai, "Paramètres · démarrage Windows, thème et IA · " + settings.ProviderName + " / " + settings.ModelName);
            tips.SetToolTip(info, "Texte transmis à " + (settings.Provider == "claude" ? "Anthropic" : "OpenAI") + " avec votre compte connecté. Aucun envoi automatique de mail.");
        }
        internal void ChangeAi(AiSettings selected, ICorrectionRunner replacement = null) {
            if (selected == null || !selected.Valid) throw new ArgumentException("Choix d’IA invalide.");
            runner.Stop(); runner.Dispose();
            settings = selected.Copy(); runner = replacement ?? new AccountRunner(aiSettings: settings);
            state.Version++; requested = -1; state.Changed = DateTime.UtcNow; InvalidateResult(); UpdateAiLabel();
            retry.Enabled = CanCorrect;
            if (manualMode) apply.Enabled = !busy && !paused;
            status.Text = "IA choisie : " + settings.ProviderName + " · " + settings.ModelName + ".";
        }
        // The choice stays pending until Enregistrer in Paramètres: Annuler keeps the current AI.
        private static AiSettings ChooseAi(Form owner, AiSettings current) {
            using (var dialog = new AiSettingsDialog(current, persist: false)) {
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.Selected : null;
            }
        }
        private void SaveAi(AiSettings selected) { AiSettingsStore.Save(selected); ChangeAi(selected); }
        private void OpenSettings() {
            if (closing || settingsOpen) return;
            settingsOpen = true;
            try {
                using (var dialog = new SettingsDialog(() => settings, ChooseAi, modePreference: modePref, startupPreference: internalMode ? new StartupPreference(new InternalStartupStore()) : null, applyAi: SaveAi, checkUpdates: testMode || internalMode ? (Func<Form, Task>)null : owner => UpdateUi.Interactive(owner, ExitApplication))) dialog.ShowDialog(this);
            } finally {
                settingsOpen = false; state.Changed = DateTime.UtcNow; defaultManual = modePref.ManualByDefault;
                retry.Enabled = CanCorrect;
                if (manualMode) apply.Enabled = !busy && !paused;
            }
        }
        internal async Task<bool> TestCorrectionDisplay() {
            await Request(false);
            return resultVersion == state.Version && apply.Enabled && result.Text.Contains("documents") && result.Text.Contains("prêts") && expanded;
        }
        internal bool TestDraftPreview() {
            using (var draft = new Draft { Key = "apercu-fictif", Text = "Bonjour,\rLes document sont prêt.\rRendez-vous à 14h.\r" }) {
                Observe(draft);
                if (manualInput.Text != DisplayText(draft.Text) || !expanded || manualInput.Height < 80 || resultVersion >= 0 || apply.Enabled || copy.Enabled) return false;
                long version = state.Version;
                Observe(draft);
                if (state.Version != version || manualInput.Text != DisplayText(draft.Text)) return false;
                draft.Text += "Merci.\r"; Observe(draft);
                if (state.Version == version || !manualInput.Text.EndsWith("Merci.\r\n") || apply.Enabled || copy.Enabled) return false;
            }
            SwitchMode(); manualInput.Text = "Texte collé, à corriger.";
            ShowDraftPreview(manualInput.Text);
            return manualInput.Text == "Texte collé, à corriger." && result.TextLength == 0 && resultVersion < 0 && !copy.Enabled && textTabs.TabPages.Count == 4;
        }
        internal async Task<bool> TestManualCorrectionDisplay() {
            SwitchMode();
            manualInput.Text = "Bonjour,\r\nLes document sont prêt pour le 28/10/2026 à 14h30.\r\nLe montant reste de 1 250,50 € HT, sous réserve de validation.\r\nPouvez-vous me confirmé votre disponibilité ?\r\nMerci";
            string original = manualInput.Text;
            await Request(false);
            string copied = null;
            CopyResult(value => copied = value);
            bool valid = manualMode && resultVersion == state.Version && copy.Enabled && copied == result.Text && copied != null && copied.Contains("documents") && copied.Contains("prêts") && copied.Contains("confirmer") && copied.Contains("sous réserve de validation") && Integrity.SameNumbers(original, copied) && manualInput.Text == original && textTabs.SelectedTab == resultPage;
            manualInput.AppendText("\r\nNouveau texte.");
            return valid && resultVersion == -1 && result.Text == copied && retry.Enabled;
        }
        private Draft CurrentDraft() { return manualMode ? new Draft { Key = "manuel", Text = manualInput.Text } : readDraft(); }
        private TextBox SelectedProposal {
            get {
                if (textTabs.SelectedTab == resultPage && correctionVersion == state.Version) return result;
                if (textTabs.SelectedTab == reformulationPage && reformulationVersion == state.Version) return reformulated;
                return null;
            }
        }
        private TextBox ShownProposal {
            get {
                if (textTabs.SelectedTab == resultPage && resultKept) return result;
                if (textTabs.SelectedTab == reformulationPage && reformulationKept) return reformulated;
                return null;
            }
        }
        private void UpdateProposalActions() {
            var proposal = SelectedProposal;
            bool valid = !busy && !paused && resultVersion == state.Version && proposal != null;
            copy.Enabled = !busy && !paused && ShownProposal != null;
            apply.Enabled = manualMode ? !busy && !paused : valid && Integrity.SameNumbers(resultSource ?? "", proposal.Text);
            // Reformuler also corrects: it is available as soon as there is text, like Corriger.
            reformulate.Enabled = !busy && !paused && !String.IsNullOrWhiteSpace(state.Text);
            adviceAsk.Enabled = !busy && !paused && !String.IsNullOrWhiteSpace(state.Text);
            foreach (var tone in toneButtons) tone.Enabled = adviceAsk.Enabled;
            adviceReformulate.Enabled = !busy && !paused && adviceVersion == state.Version && adviceBox.TextLength > 0;
        }
        private void SetOriginal(string text) {
            string displayed = DisplayText(text);
            if (manualInput.Text == displayed) return;
            updatingInput = true;
            try { manualInput.Text = displayed; } finally { updatingInput = false; }
        }
        private void SwitchMode() {
            if (busy) runner.Stop();
            manualMode = !manualMode; awayArmed = false;
            context.Text = manualMode ? "MODE MANUEL" : "OUTLOOK";
            manualInput.ReadOnly = !manualMode;
            SetOriginal(manualMode ? manualDraft : "");
            state.Observe(manualMode ? "manuel" : "", manualMode ? manualInput.Text : "", DateTime.UtcNow);
            requested = -1; InvalidateResult();
            manual.Text = manualMode ? "OUTLOOK" : "Manuel"; apply.Text = manualMode ? "Coller" : "Appliquer";
            if (manualMode) {
                textTabs.SelectedTab = inputPage; SetExpanded(true);
                apply.Enabled = !busy && !paused; retry.Enabled = CanCorrect;
                status.Text = "Écrivez ou collez votre texte, puis Corriger.";
                manualInput.Focus();
            } else {
                textTabs.SelectedTab = inputPage; SetExpanded(false);
                apply.Enabled = retry.Enabled = false;
                status.Text = "Mode OUTLOOK · attente du brouillon.";
            }
            reformulate.Enabled = false;
            Palette.Apply(this);
        }
        internal void ImportSelection(string text) {
            if (closing || IsDisposed || !SelectionAccess.Valid(text)) return;
            if (busy) runner.Stop();
            if (!manualMode) SwitchMode();
            manualInput.Text = DisplayText(text); manualDraft = manualInput.Text;
            state.Observe("manuel", manualInput.Text, DateTime.UtcNow); state.Version++; requested = -1;
            InvalidateResult(); textTabs.SelectedTab = inputPage; SetExpanded(true);
            RestoreFromTray(); Activate(); manualInput.Focus(); manualInput.SelectionStart = 0; manualInput.SelectionLength = 0;
            retry.Enabled = CanCorrect; UpdateProposalActions();
            status.Text = "Sélection reçue · choisissez Corriger ou Reformuler. Alt + K depuis une autre application.";
        }
        private void PasteManual() {
            if (!manualMode || busy || paused) return;
            textTabs.SelectedTab = inputPage; SetExpanded(true);
            try {
                if (Clipboard.ContainsText()) { manualInput.Focus(); manualInput.Paste(); }
                else status.Text = "Copiez d’abord le texte de votre mail.";
            } catch (ExternalException) { status.Text = "Presse-papiers occupé · réessayez."; }
        }
        private void CopyResult(Action<string> writer) {
            var shown = ShownProposal;
            if (busy || closing || paused || shown == null) return;
            bool upToDate;
            using (var current = CurrentDraft()) {
                // A proposal is never copied for another mail.
                if (current == null || current.Key != keptKey) {
                    InvalidateResult(); status.Text = "Le brouillon a changé. Relancez la correction avant de copier."; return;
                }
                upToDate = SelectedProposal == shown && state.Matches(resultVersion, current.Key, current.Text) && current.Key == resultKey && current.Text == resultSource;
            }
            try {
                if (writer != null) writer(shown.Text); else Clipboard.SetText(shown.Text);
                status.Text = upToDate ? "Correction copiée · collez-la dans votre mail." : "Proposition précédente copiée · sans vos dernières modifications.";
            } catch (ExternalException) { status.Text = "Copie impossible · sélectionnez le texte puis Ctrl+C."; }
        }
        // Word uses vertical tabs for manual (Shift+Enter) line breaks. WinForms
        // TextBox shows those as missing glyphs unless converted for display.
        // Keep the raw draft untouched for identity and stale-result validation.
        private static string DisplayText(string text) { return text.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\v", "\n").Replace("\u2028", "\n").Replace("\u2029", "\n").Replace("\n", "\r\n"); }
        private static bool SameWording(string left, string right) {
            return Regex.Replace(left.Trim(), @"\s+", " ") == Regex.Replace(right.Trim(), @"\s+", " ");
        }
        private void ShowDraftPreview(string text) {
            if (!manualMode) SetOriginal(text);
            // A kept proposal stays in view while the draft is edited.
            if (resultVersion < 0 && ShownProposal == null && textTabs.SelectedTab != inputPage) textTabs.SelectedTab = inputPage;
            SetExpanded(true);
            UpdateProposalActions();
        }
        private void SetExpanded(bool value) {
            expand.Enabled = manualMode || !String.IsNullOrWhiteSpace(state.Text);
            if (value && !expand.Enabled) value = false;
            if (expanded == value) return;
            bool opening = value && !expanded;
            expanded = value;
            int bottom = Bottom;
            textTabs.Visible = value; expand.Text = value ? "Réduire" : "Aperçu";
            float scale = CurrentAutoScaleDimensions.Height / 96f;
            if (!value) expandedClientSize = ClientSize;
            MinimumSize = SizeFromClientSize(new Size((int)(480 * scale), (int)((value ? 300 : 160) * scale)));
            if (opening) ClientSize = new Size(ClientSize.Width, Math.Max(ClientSize.Height, Math.Max(expandedClientSize.Height, (int)Math.Round(350 * scale))));
            else ClientSize = new Size(ClientSize.Width, (int)Math.Round(160 * scale));
            var area = Screen.FromControl(this).WorkingArea;
            Top = Math.Max(area.Top, Math.Min(bottom, area.Bottom) - Height);
        }
        // keepProposals: an edit of the same draft keeps the Corrigé and Reformulé texts on screen
        // (still copyable); only Appliquer needs a new correction of the current text.
        private void InvalidateResult(bool keepPreview = false, bool keepProposals = false) {
            resultVersion = -1; resultKey = resultSource = null;
            correctionVersion = reformulationVersion = adviceVersion = -1;
            apply.Enabled = reformulate.Enabled = copy.Enabled = adviceReformulate.Enabled = false;
            if (keepProposals && (resultKept || reformulationKept)) { UpdateProposalActions(); return; }
            resultKept = reformulationKept = false; keptKey = null;
            if (textTabs.SelectedTab != inputPage) textTabs.SelectedTab = inputPage;
            // A new keystroke invalidates the correction without hiding or clearing
            // the live preview. ShowDraftPreview replaces its content directly.
            if (result.TextLength > 0) result.Clear();
            if (reformulated.TextLength > 0) reformulated.Clear();
            if (adviceBox.TextLength > 0) adviceBox.Clear();
            if (keepPreview) return;
            SetExpanded(manualMode);
        }
        private void Observe(Draft draft) {
            if (!manualMode) {
                if (!draftNumbers.ContainsKey(draft.Key)) draftNumbers[draft.Key] = draftNumbers.Count + 1;
                context.Text = "OUTLOOK · MAIL " + draftNumbers[draft.Key];
                tips.SetToolTip(context, "Repère du brouillon dans cette session. Seul le mail de la fenêtre active est corrigé.");
            }
            string previousKey = state.Key;
            if (state.Observe(draft.Key, draft.Text, DateTime.UtcNow)) {
                bool preview = !manualMode && !String.IsNullOrWhiteSpace(draft.Text);
                InvalidateResult(preview, draft.Key == previousKey && !String.IsNullOrWhiteSpace(draft.Text));
                if (busy) runner.Stop();
                if (preview) ShowDraftPreview(draft.Text);
                else status.Text = String.IsNullOrWhiteSpace(draft.Text) ? "Brouillon vide · écrivez votre texte." : "Texte manuel · cliquez sur Corriger.";
                if (!manualMode) Recall();
            }
        }
        private static string HandledKey(string key, string text) { return key + "\u0001" + text; }
        private void Remember(string key, string text, string correction, bool replace) {
            if (manualMode || String.IsNullOrWhiteSpace(text) || !replace && handled.ContainsKey(HandledKey(key, text))) return;
            if (handled.Count >= 100) handled.Clear();
            handled[HandledKey(key, text)] = correction;
        }
        // A text seen again never triggers the automatic correction; Corriger stays available by click.
        private void Recall() {
            string correction;
            if (!handled.TryGetValue(HandledKey(state.Key, state.Text), out correction)) return;
            long version = state.Version; requested = version;
            if (correction != null) {
                result.Text = DisplayText(correction); correctionVersion = resultVersion = correctedVersion = version; resultKey = state.Key; resultSource = state.Text; RefreshChanges();
                resultKept = true; keptKey = state.Key;
                textTabs.SelectedTab = resultPage; SetExpanded(true);
                status.Text = "Texte déjà corrigé · correction précédente réaffichée, sans nouvelle demande.";
            } else status.Text = "Texte déjà traité par Korecto · Corriger reste possible.";
            retry.Enabled = CanCorrect; UpdateProposalActions();
        }
        // Reading Outlook never calls the AI: corrections start only from a click (Corriger, Reformuler).
        private Task Tick() { RefreshPreview(); return Task.FromResult(true); }
        private bool RefreshPreview() {
            if (closing || paused || settingsOpen) return false;
            // Manuel is a detour from the default Outlook mode: back once Outlook was left and a draft is in front again.
            if (manualMode && !defaultManual && !busy && ReturnedToOutlook()) SwitchMode();
            if (hiddenByUser && manualMode) return false;
            if (manualMode) {
                retry.Enabled = CanCorrect; UpdateProposalActions();
                if (busy) status.Text = requestTitle + " · " + (int)(DateTime.UtcNow - requestStarted).TotalSeconds + " s\n" + runner.Progress;
                return false;
            }
            bool foreground = outlookForeground();
            if (!foreground && (!Visible || !ContainsFocus)) {
                // Default Manuel: the Outlook mode lasts only while an Outlook draft is in front.
                if (defaultManual) SwitchMode(); else if (detected) Hide();
                return false;
            }
            using (var draft = readDraft()) {
                // A ribbon menu or dialog hides the draft from Outlook for a moment: keep Korecto as it is.
                if (draft == null && foreground && outlookPopup()) return false;
                if (hiddenByUser) {
                    if (!foreground || draft == null || draft.Key == state.Key && draft.Text == state.Text) return false;
                    hiddenByUser = false; ShowInTaskbar = true;
                }
                if (draft == null) {
                    bool left = detected;
                    if (detected || state.Key.Length > 0) { runner.Stop(); state.Observe("", "", DateTime.UtcNow); InvalidateResult(); SetOriginal(""); detected = false; }
                    SetExpanded(false);
                    status.Text = "Ouvrez un brouillon Outlook ou passez en Manuel.";
                    // Draft left, closed or sent: Korecto withdraws to the notification area and
                    // returns with the next draft (Manuel stays in the notification menu).
                    if (left && defaultManual) { SwitchMode(); return false; }
                    if (left && Visible) Hide();
                    return false;
                }
                detected = true;
                if (!Visible) { ShowInTaskbar = true; WindowState = FormWindowState.Normal; Show(); }
                Observe(draft);
                retry.Enabled = CanCorrect;
                UpdateProposalActions();
                if (busy) status.Text = requestTitle + " · " + (int)(DateTime.UtcNow - requestStarted).TotalSeconds + " s\n" + runner.Progress;
                if (!busy && resultVersion < 0 && !String.IsNullOrWhiteSpace(state.Text) && manualInput.Text != DisplayText(state.Text)) ShowDraftPreview(state.Text);
                if (!busy && requested != state.Version && !String.IsNullOrWhiteSpace(state.Text)) {
                    status.Text = resultKept || reformulationKept ? "Texte modifié · proposition précédente gardée. Corriger pour l’actualiser." : "Texte en direct · cliquez sur Corriger quand vous avez fini.";
                }
            }
            return true;
        }
        // Words changed since the original text, shown with a soft background in Corrigé (the text itself is untouched).
        private void RefreshChanges() {
            if (result.TextLength == 0 || String.IsNullOrEmpty(resultSource)) return;
            // A RichTextBox counts line breaks as one character.
            string original = DisplayText(resultSource).Replace("\r\n", "\n"), shown = result.Text.Replace("\r\n", "\n");
            var handle = resultView.Handle;
            if (resultView.Text != shown) resultView.Text = shown;
            resultView.SelectAll(); resultView.SelectionBackColor = resultView.BackColor;
            Color color = Palette.IsDark ? Color.FromArgb(52, 96, 70) : Color.FromArgb(214, 237, 207);
            int delta = 0;
            foreach (var edit in Differences(original, shown)) {
                int length = edit.Item3.Length, start = edit.Item1 + delta; delta += length - (edit.Item2 - edit.Item1);
                if (length == 0 || start + length > resultView.TextLength) continue;
                resultView.Select(start, length); resultView.SelectionBackColor = color;
            }
            resultView.Select(0, 0);
        }
        // A vertical scrollbar only when the text does not fit.
        private static void FitScrollbar(TextBox box) {
            if (box.TextLength == 0 && box.ScrollBars == ScrollBars.None) return;
            int lines = box.GetLineFromCharIndex(box.TextLength) + 1;
            var wanted = lines * box.Font.Height > box.ClientSize.Height ? ScrollBars.Vertical : ScrollBars.None;
            if (wanted == box.ScrollBars) return;
            int start = box.SelectionStart, length = box.SelectionLength;
            box.ScrollBars = wanted; box.Select(start, length);
            if (box.Focused) box.ScrollToCaret();
        }
        private void EnsureHeight(int logicalHeight) {
            int needed = (int)Math.Round(logicalHeight * CurrentAutoScaleDimensions.Height / 96f);
            if (ClientSize.Height >= needed) return;
            int bottom = Bottom; ClientSize = new Size(ClientSize.Width, needed);
            var area = Screen.FromControl(this).WorkingArea; Top = Math.Max(area.Top, Math.Min(bottom, area.Bottom) - Height);
        }
        private bool ReturnedToOutlook() {
            if (!outlookForeground()) { awayArmed = true; return false; }
            if (!awayArmed) return false;
            using (var draft = readDraft()) return draft != null;
        }
        private void TogglePause() {
            paused = !paused; pause.Text = paused ? "▶" : "Ⅱ"; pause.AccessibleName = paused ? "Reprendre" : "Mettre en pause"; tips.SetToolTip(pause, pause.AccessibleName);
            if (paused) { runner.Stop(); state.Version++; InvalidateResult(); retry.Enabled = false; status.Text = "En pause · aucune nouvelle lecture ni transmission."; }
            else { requested = -1; state.Changed = DateTime.UtcNow; retry.Enabled = !busy && !String.IsNullOrWhiteSpace(state.Text); if (manualMode) apply.Enabled = !busy; status.Text = manualMode ? "Texte manuel · cliquez sur Corriger." : "Aperçu repris · cliquez sur Corriger."; }
        }
        private Task Request(bool fluent) { return RequestCore(fluent, null); }
        // advice: the answer of the Avis tab, which the reformulation must follow.
        private async Task RequestCore(bool fluent, string advice) {
            if (busy || closing || paused || settingsOpen) return;
            using (var draft = CurrentDraft()) {
                if (draft == null) { status.Text = "Ouvrez le brouillon dans sa propre fenêtre OUTLOOK."; return; }
                Observe(draft);
                if (String.IsNullOrWhiteSpace(state.Text)) return;
                if (!fluent && correctedVersion == state.Version) return;
                if (state.Text.Length > 30000) { requested = state.Version; status.Text = "Brouillon trop long (30 000 caractères maximum)."; return; }
                long version = state.Version; string key = state.Key, text = state.Text;
                string previousProposal = resultVersion == version ? (SelectedProposal ?? (reformulationVersion == version ? reformulated : result)).Text : null;
                TabPage previousPage = textTabs.SelectedTab;
                requested = version; busy = true; apply.Enabled = reformulate.Enabled = retry.Enabled = copy.Enabled = adviceAsk.Enabled = adviceReformulate.Enabled = false; foreach (var tone in toneButtons) tone.Enabled = false;
                requestStarted = DateTime.UtcNow; requestTitle = fluent ? "Reformulation" : "Correction stricte";
                ShowDraftPreview(text);
                status.Text = advice != null ? "Reformulation selon l’avis en cours…" : fluent ? "Reformulation en cours…" : "Correction stricte en cours…";
                try {
                    string corrected = await runner.Correct(text, fluent, advice);
                    if (closing || paused || state.Version != version) return;
                    if (String.IsNullOrWhiteSpace(corrected)) throw new InvalidOperationException(settings.ProviderName + " n’a pas renvoyé de " + (fluent ? "reformulation" : "correction") + " exploitable. Cliquez sur " + (fluent ? "Reformuler." : "Corriger."));
                    using (var current = CurrentDraft()) {
                        if (current == null || !state.Matches(version, current.Key, current.Text)) { InvalidateResult(true, current != null && current.Key == keptKey); status.Text = "Le brouillon a changé · nouvelle correction nécessaire."; return; }
                    }
                    bool unchanged = fluent && (SameWording(corrected, text) || previousProposal != null && SameWording(corrected, previousProposal) || correctionVersion == version && SameWording(corrected, result.Text) || reformulationVersion == version && SameWording(corrected, reformulated.Text));
                    // Preserve a valid correction if the model only repeats the source.
                    if (unchanged && reformulationVersion == version) corrected = reformulated.Text;
                    else if (unchanged && previousProposal != null) corrected = previousProposal;
                    if (fluent) { reformulated.Text = DisplayText(corrected); reformulationVersion = version; reformulationKept = true; }
                    else { result.Text = DisplayText(corrected); correctionVersion = version; resultKept = true; }
                    keptKey = key;
                    resultVersion = version; resultKey = key; resultSource = text;
                    if (!fluent) RefreshChanges();
                    // A reformulation that alters numbers must not cancel a valid strict correction.
                    if (Integrity.SameNumbers(text, corrected)) correctedVersion = version;
                    else if (!fluent) correctedVersion = -1;
                    if (!fluent && correctedVersion == version) Remember(key, text, corrected, true);
                    textTabs.SelectedTab = fluent ? reformulationPage : resultPage;
                    SetExpanded(true);
                    status.Text = fluent ? unchanged ? "Aucune reformulation proposée · texte inchangé." : "Reformulation proposée · vérifiez avant d’appliquer." : "Correction stricte proposée · vérifiez avant d’appliquer.";
                    if (!Integrity.SameNumbers(text, corrected)) status.Text = "Les chiffres ont changé · vérifiez la proposition.";
                } catch (Exception ex) {
                    if (!closing && !paused && state.Version == version) {
                        if (fluent && previousProposal != null) {
                            resultVersion = version; resultKey = key; resultSource = text;
                            textTabs.SelectedTab = previousPage;
                            status.Text = "Reformulation indisponible · proposition précédente conservée. " + ex.Message;
                        } else {
                            resultVersion = -1;
                            string error = ex is InvalidOperationException ? ex.Message : (fluent ? "Reformulation" : "Correction") + " interrompue ou indisponible. Cliquez sur " + (fluent ? "Reformuler" : "Corriger") + " pour réessayer.";
                            // An earlier proposal kept in the tab is never replaced by an error message.
                            if (fluent ? reformulationKept : resultKept) status.Text = (fluent ? "Reformulation" : "Correction") + " indisponible · proposition précédente conservée. " + error;
                            else {
                                status.Text = (fluent ? "Reformulation" : "Correction") + " indisponible · cliquez sur " + (fluent ? "Reformuler" : "Corriger") + ".";
                                (fluent ? reformulated : result).Text = error;
                                textTabs.SelectedTab = fluent ? reformulationPage : resultPage;
                            }
                        }
                        SetExpanded(true);
                    }
                } finally { busy = false; if (!closing) { retry.Enabled = CanCorrect; UpdateProposalActions(); } }
            }
        }
        private async Task RequestAdvice() {
            if (busy || closing || paused || settingsOpen) return;
            using (var draft = CurrentDraft()) {
                if (draft == null) { status.Text = "Ouvrez le brouillon dans sa propre fenêtre OUTLOOK."; return; }
                Observe(draft);
                if (String.IsNullOrWhiteSpace(state.Text)) return;
                if (state.Text.Length > 30000) { status.Text = "Brouillon trop long (30 000 caractères maximum)."; return; }
                long version = state.Version; string text = state.Text;
                string question = adviceQuestion.Text.Trim(); if (question.Length == 0) question = "Donne ton avis sur ce mail.";
                busy = true; apply.Enabled = reformulate.Enabled = retry.Enabled = copy.Enabled = adviceAsk.Enabled = adviceReformulate.Enabled = false; foreach (var tone in toneButtons) tone.Enabled = false;
                requestStarted = DateTime.UtcNow; requestTitle = "Avis";
                status.Text = "Avis en cours…";
                try {
                    string answer = await runner.Advise(text, question);
                    if (closing || paused || state.Version != version) return;
                    if (String.IsNullOrWhiteSpace(answer)) throw new InvalidOperationException(settings.ProviderName + " n’a pas renvoyé d’avis exploitable. Cliquez sur Demander l’avis.");
                    using (var current = CurrentDraft()) {
                        if (current == null || !state.Matches(version, current.Key, current.Text)) { status.Text = "Le brouillon a changé · demandez l’avis à nouveau."; return; }
                    }
                    adviceBox.Text = DisplayText(answer); adviceVersion = version;
                    status.Text = "Avis reçu · « Reformuler selon cet avis » pour le suivre.";
                } catch (Exception ex) {
                    if (!closing && !paused && state.Version == version) status.Text = "Avis indisponible. " + (ex is InvalidOperationException ? ex.Message : "Cliquez sur Demander l’avis pour réessayer.");
                } finally { busy = false; if (!closing) { retry.Enabled = CanCorrect; UpdateProposalActions(); } }
            }
        }
        private void Apply() {
            var proposal = SelectedProposal;
            if (busy || paused || resultVersion != state.Version || proposal == null || !Integrity.SameNumbers(resultSource ?? "", proposal.Text)) return;
            using (var draft = readDraft()) {
                if (draft == null || !state.Matches(resultVersion, draft.Key, draft.Text) || draft.Key != resultKey || draft.Text != resultSource) {
                    InvalidateResult(true, draft != null && draft.Key == keptKey); status.Text = "Le brouillon a changé. Aucune modification appliquée."; return;
                }
                try {
                    // Replace only the captured author range, never subject, recipients or attachments.
                    string corrected = proposal.Text.Replace("\r\n", "\r").Replace("\n", "\r");
                    if (draft.Text.EndsWith("\r") && !corrected.EndsWith("\r")) corrected += "\r";
                    int linksKept = ReplaceDraftText(draft, corrected);
                    // Undoing an applied reformulation restores a source with no strict correction: still no new request.
                    Remember(draft.Key, draft.Text, null, false);
                    long oldVersion = state.Version;
                    using (var current = readDraft()) { if (current != null) state.Observe(current.Key, current.Text, DateTime.UtcNow); }
                    if (correctionVersion == oldVersion) correctionVersion = state.Version;
                    if (reformulationVersion == oldVersion) reformulationVersion = state.Version;
                    resultVersion = state.Version; resultKey = state.Key; resultSource = state.Text;
                    correctedVersion = state.Version;
                    Remember(state.Key, state.Text, null, false);
                    requested = state.Version; UpdateProposalActions(); status.Text = linksKept > 0 ? "Proposition appliquée · liens et mentions gardés tels quels." : "Proposition appliquée · original et variantes conservés ici.";
                } catch (Exception ex) { status.Text = ex.Message == ComplexDraft ? ComplexDraft : "Application impossible. Le brouillon reste à vérifier dans OUTLOOK."; }
            }
        }
        internal static void TextChange(string original, string replacement, out int start, out int end, out string inserted) {
            start = 0;
            while (start < original.Length && start < replacement.Length && original[start] == replacement[start]) start++;
            if (start > 0 && start < original.Length && Char.IsLowSurrogate(original[start])) start--;
            int tail = 0;
            while (tail < original.Length - start && tail < replacement.Length - start && original[original.Length - 1 - tail] == replacement[replacement.Length - 1 - tail]) tail++;
            if (tail > 0 && tail < original.Length && Char.IsLowSurrogate(original[original.Length - tail])) tail--;
            end = original.Length - tail;
            inserted = replacement.Substring(start, replacement.Length - tail - start);
        }
        internal const string ComplexDraft = "Image, tableau ou lien inhabituel dans ce mail · Appliquer bloqué pour ne pas l’abîmer. Utilisez Copier.";
        // Returns the number of changes left out because they fell inside a link.
        internal static int ReplaceDraftText(Draft draft, string corrected) {
            if (draft.Document == null) { ((dynamic)draft.Range).Text = corrected; return 0; } // isolated test draft
            int origin = (int)((dynamic)draft.Range).Start;
            // Links and @mentions are Word fields: their hidden code counts in Word positions, not in Range.Text.
            var links = Links(draft, origin);
            int hidden = 0; foreach (var link in links) hidden += link[2];
            if ((int)((dynamic)draft.Range).End - origin - hidden != draft.Text.Length) throw new InvalidOperationException(ComplexDraft);
            int skipped; var edits = PlanEdits(draft.Text, corrected, links, out skipped);
            if (edits.Count == 0) return skipped;
            // One Ctrl+Z in Outlook must undo the whole proposal, not one paragraph at a time
            // (intermediate texts would look new and trigger a correction). Word 2010 and later.
            object word = null, undo = null;
            try { word = ((dynamic)draft.Document).Application; undo = ((dynamic)word).UndoRecord; ((dynamic)undo).StartCustomRecord("Korecto"); }
            catch (COMException) { Native.Release(undo); undo = null; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { Native.Release(undo); undo = null; }
            finally { Native.Release(word); }
            try {
                // Every change is checked before the first write: a refusal never leaves a half-applied draft.
                for (int i = edits.Count - 1; i >= 0; i--) {
                    object changedRange = null;
                    try { changedRange = ((dynamic)draft.Document).Range(WordPosition(origin, links, edits[i].Item1), WordPosition(origin, links, edits[i].Item2)); ((dynamic)changedRange).Text = edits[i].Item3; }
                    finally { Native.Release(changedRange); }
                }
            } finally {
                if (undo != null) { try { ((dynamic)undo).EndCustomRecord(); } catch (COMException) { } Native.Release(undo); }
            }
            return skipped;
        }
        // Each field as { label start, label end } in Range.Text offsets, plus the length Word hides around the label.
        private static List<int[]> Links(Draft draft, int origin) {
            var links = new List<int[]>(); int hidden = 0, previousEnd = 0;
            object fields = null;
            try {
                fields = ((dynamic)draft.Range).Fields; int count = (int)((dynamic)fields).Count;
                for (int i = 1; i <= count; i++) {
                    object field = null, code = null, label = null;
                    try {
                        field = ((dynamic)fields).Item(i); code = ((dynamic)field).Code; label = ((dynamic)field).Result;
                        int begin = (int)((dynamic)code).Start - 1 - origin, separator = (int)((dynamic)label).Start - origin, close = (int)((dynamic)label).End - origin;
                        int start = begin - hidden, end = start + close - separator;
                        if (begin < previousEnd || separator <= begin || close < separator || start < 0 || end > draft.Text.Length || draft.Text.Substring(start, end - start) != Convert.ToString(((dynamic)label).Text)) throw new InvalidOperationException(ComplexDraft);
                        links.Add(new [] { start, end, separator - begin + 1 });
                        hidden += separator - begin + 1; previousEnd = close + 1;
                    } finally { Native.Release(label); Native.Release(code); Native.Release(field); }
                }
            } finally { Native.Release(fields); }
            return links;
        }
        // A position at a label's start stays before its field; one at its end goes after the field.
        private static int WordPosition(int origin, List<int[]> links, int offset) {
            int position = origin + offset;
            foreach (var link in links) if (link[1] <= offset) position += link[2];
            return position;
        }
        // Changes word by word; a change inside a link is left out, the link stays exactly as it is.
        internal static List<Tuple<int, int, string>> PlanEdits(string original, string replacement, List<int[]> links, out int skipped) {
            var kept = new List<Tuple<int, int, string>>(); skipped = 0;
            foreach (var edit in Differences(original, replacement)) {
                if (links.Exists(link => edit.Item1 < link[1] && edit.Item2 > link[0] || edit.Item1 == edit.Item2 && link[0] < edit.Item1 && edit.Item1 < link[1])) { skipped++; continue; }
                var last = kept.Count == 0 ? null : kept[kept.Count - 1];
                // Neighbouring changes of one paragraph form one write, unless a link lies between them.
                if (last != null && original.IndexOf('\r', last.Item2, edit.Item1 - last.Item2) < 0 && !links.Exists(link => link[0] < edit.Item1 && link[1] > last.Item2))
                    kept[kept.Count - 1] = Tuple.Create(last.Item1, edit.Item2, last.Item3 + original.Substring(last.Item2, edit.Item1 - last.Item2) + edit.Item3);
                else kept.Add(edit);
            }
            var edits = new List<Tuple<int, int, string>>();
            foreach (var edit in kept) {
                int start, end; string inserted; TextChange(original.Substring(edit.Item1, edit.Item2 - edit.Item1), edit.Item3, out start, out end, out inserted);
                if (start == end && inserted.Length == 0) continue;
                // Word marks images, tables and fields with control characters; rewriting them would delete the object.
                if (Regex.IsMatch(original.Substring(edit.Item1 + start, end - start) + inserted, @"[\x00-\x08\x0E-\x1F]")) throw new InvalidOperationException(ComplexDraft);
                edits.Add(Tuple.Create(edit.Item1 + start, edit.Item1 + end, inserted));
            }
            return edits;
        }
        private static readonly Regex Token = new Regex(@"\r|[^\S\r]+|\w+|[\uD800-\uDBFF][\uDC00-\uDFFF]|[^\w\s]");
        // Changed runs of words between the two texts (longest common sequence of words).
        private static List<Tuple<int, int, string>> Differences(string original, string replacement) {
            var result = new List<Tuple<int, int, string>>();
            int start, end; string inserted; TextChange(original, replacement, out start, out end, out inserted);
            if (start == end && inserted.Length == 0) return result;
            var a = Token.Matches(original.Substring(start, end - start)); var b = Token.Matches(inserted);
            if ((long)a.Count * b.Count > 4000000) { result.Add(Tuple.Create(start, end, inserted)); return result; }
            var common = new int[a.Count + 1, b.Count + 1];
            for (int i = a.Count - 1; i >= 0; i--) for (int j = b.Count - 1; j >= 0; j--)
                common[i, j] = a[i].Value == b[j].Value ? common[i + 1, j + 1] + 1 : Math.Max(common[i + 1, j], common[i, j + 1]);
            int x = 0, y = 0, fromA = 0, fromB = 0;
            while (x < a.Count || y < b.Count) {
                if (x < a.Count && y < b.Count && a[x].Value == b[y].Value) {
                    if (a[x].Index > fromA || b[y].Index > fromB) result.Add(Tuple.Create(start + fromA, start + a[x].Index, inserted.Substring(fromB, b[y].Index - fromB)));
                    fromA = a[x].Index + a[x].Length; fromB = b[y].Index + b[y].Length; x++; y++;
                } else if (y >= b.Count || x < a.Count && common[x + 1, y] >= common[x, y + 1]) x++;
                else y++;
            }
            if (end - start > fromA || inserted.Length > fromB) result.Add(Tuple.Create(start + fromA, end, inserted.Substring(fromB)));
            return result;
        }
    }
    internal static class Program {
        internal static string TestModelLine(AiSettings ai) {
            return "IA et modèle demandés : " + ai.ProviderName + " / " + ai.Model + (ai.Provider == "chatgpt" ? ", effort faible." : ".");
        }
        [STAThread] private static void Main(string[] args) {
            bool internalMode = Array.IndexOf(args, "--internal") >= 0;
            AiSettings internalAi = null;
            if (internalMode) {
                internalAi = AiSettingsStore.Load();
                Environment.SetEnvironmentVariable("KORECTO_DONNEES", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DONNEES_TEST"));
            }
            bool startupLaunch = Array.IndexOf(args, "--startup") >= 0;
            string root = AppDomain.CurrentDomain.BaseDirectory;
            if (args.Length > 0 && (args[0] == "--test-ui-account" || args[0] == "--test-manual-account")) {
                bool manualTest = args[0] == "--test-manual-account";
                Draft.SyntheticTest = !manualTest;
                Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
                var ai = AiSettings.Default();
                using (var popup = new Popup(true, aiSettings: ai)) {
                    popup.Opacity = 0; popup.ShowInTaskbar = false;
                    popup.Load += async (sender, e) => {
                        var watch = Stopwatch.StartNew(); bool valid = false;
                        try { valid = manualTest ? await popup.TestManualCorrectionDisplay() : await popup.TestCorrectionDisplay(); }
                        finally {
                            string checks = manualTest ? "Saisie manuelle, réception et affichage, correction des accords et de l’infinitif, chiffres et réserve conservés, texte original conservé, commande Copier et invalidation après modification : " : "Réponse reçue, texte placé dans la fenêtre, Appliquer actif et proposition dépliée : ";
                            File.WriteAllText(Path.Combine(root, manualTest ? "TEST_MODE_MANUEL.txt" : "TEST_INTERFACE.txt"), "Test de bout en bout sur texte fictif : " + (valid ? "réussi" : "échec") + "\r\n" + TestModelLine(ai) + "\r\n" + checks + valid + "\r\nDurée : " + Math.Round(watch.Elapsed.TotalSeconds, 1) + " secondes.\r\nAucun mail réel lu ou transmis. Presse-papiers réel inchangé par le test.", Encoding.UTF8);
                            Environment.ExitCode = valid ? 0 : 1; popup.ExitApplication();
                        }
                    };
                    Application.Run(popup);
                } return;
            }
            if (args.Length > 0 && args[0] == "--diagnostic-outlook") {
                var report = new StringBuilder(); string previousKey = null, previousText = null;
                var debounce = new Debounce(); bool stable = true;
                for (int i = 0; i < 5; i++) {
                    using (var draft = Draft.Read()) {
                        if (draft == null) { stable = false; report.AppendLine("Lecture " + i + " : brouillon non accessible"); }
                        else {
                            if (i > 0 && (draft.Key != previousKey || draft.Text != previousText)) stable = false;
                            report.AppendLine("Lecture " + i + " : accessible ; même identifiant=" + (draft.Key == previousKey) + "; même texte=" + (draft.Text == previousText));
                            debounce.Observe(draft.Key, draft.Text, DateTime.UtcNow);
                            previousKey = draft.Key; previousText = draft.Text;
                        }
                    }
                    Thread.Sleep(600);
                }
                report.AppendLine("Détection stable : " + stable);
                report.AppendLine("Correction uniquement sur clic (Corriger) : la lecture d’Outlook n’envoie rien à l’IA.");
                File.WriteAllText(Path.Combine(root, "DIAGNOSTIC_OUTLOOK.txt"), report.ToString(), Encoding.UTF8); return;
            }
            if (args.Length > 0 && args[0] == "--self-test") {
                var now = DateTime.UtcNow; var d = new Debounce(); int count = 0;
                Action<bool> check = ok => { if (!ok) throw new Exception("Test " + (count + 1) + " échoué"); count++; };
                check(d.Observe("a", "Bonjour", now));
                long v = d.Version; check(!d.Observe("a", "Bonjour", now.AddSeconds(11))); check(d.Matches(v, "a", "Bonjour"));
                check(d.Observe("a", "Bonjour !", now.AddSeconds(12))); check(!d.Matches(v, "a", "Bonjour"));
                check(d.Observe("b", "Bonjour !", now.AddSeconds(22))); check(!d.Matches(d.Version, "a", "Bonjour !"));
                d.Observe("b", " ", now); check(String.IsNullOrWhiteSpace(d.Text));
                check(Integrity.SameNumbers("1 250,50 € avant le 26/10/2026", "1 250,50 € avant le 26/10/2026."));
                check(!Integrity.SameNumbers("1 250,50 €", "1 200,50 €"));
                check(!Integrity.SameNumbers("Sous réserve de validation.", "Validation sous 10 jours."));
                check(!Integrity.SameNumbers("26/10/2026", "27/10/2026"));
                check(!Integrity.SameNumbers("100 %", "100" + ",5 %"));
                var parser = new JavaScriptSerializer(); parser.DeserializeObject(AppAssets.Read(root, "response.schema.json"));
                File.WriteAllText(Path.Combine(root, "VERIFICATION.txt"), count + " contrôles réussis : modification, changement de fenêtre, résultat périmé, texte vide, intégrité des chiffres et dates.\r\nSchéma JSON valide.\r\nCodex trouvé : " + (AccountRunner.FindCodex() != null) + "\r\n", Encoding.UTF8); return;
            }
            if (args.Length > 0 && args[0] == "--test-window") {
                Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
                using (var popup = new Popup(true)) {
                    // Let WinForms complete layout, just as it does on a real launch.
                    popup.Opacity = 0; popup.ShowInTaskbar = false; popup.Show(); popup.PerformLayout();
                    bool valid = popup.Controls.Count == 1 && popup.TopMost && popup.ClientSize.Width >= 540 && popup.TestDraftPreview();
                    popup.ExitApplication();
                    File.WriteAllText(Path.Combine(root, "TEST_FENETRE.txt"), valid ? "Initialisation et fermeture WinForms réussies. Aperçu du texte pendant l’attente vérifié sur texte fictif : mise à jour, distinction de la correction, boutons Appliquer/Copier désactivés, texte manuel conservé. Aucun test réseau, visuel ou OUTLOOK réel effectué." : "Échec initialisation ou aperçu.", Encoding.UTF8);
                    Environment.ExitCode = valid ? 0 : 1;
                } return;
            }
            if (args.Length > 0 && args[0] == "--test-connection") {
                using (var runner = new AccountRunner()) {
                    try { string corrected = runner.Correct("Bonjour,\nLes document sont prêt.\nMerci", false).GetAwaiter().GetResult();
                        File.WriteAllText(Path.Combine(root, "TEST_CONNEXION.txt"), "Test fictif :\r\n" + corrected, Encoding.UTF8); }
                    catch (Exception ex) { File.WriteAllText(Path.Combine(root, "TEST_CONNEXION.txt"), "Échec : " + ex.Message, Encoding.UTF8); Environment.ExitCode = 1; }
                } return;
            }
            if (args.Length > 0 && args[0] == "--desinstaller") { SelfInstaller.RunUninstallUi(); return; }
            if (args.Length > 0 && args[0] == "--appliquer-maj") { UpdateUi.RunApply(); return; }
            // A downloaded Korecto.exe offers to install itself; the installed copy and startup launches never ask.
            if (!startupLaunch && !internalMode && SelfInstaller.OfferInstall(Application.ExecutablePath)) return;
            bool created;
            using (var mutex = new Mutex(true, internalMode ? "Local\\KorectoInterne" : "Local\\CorrecteurMailChatGPT", out created)) {
                if (!created) { if (!startupLaunch) MessageBox.Show("Korecto est déjà lancé. Retrouvez son icône près de l’horloge."); return; }
                Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
                if (!internalMode) AppFolders.MigrateLegacy(Path.Combine(root, "PARAMETRES_IA"), AppFolders.Data);
                Palette.Use(AppearanceStore.LoadDark());
                AiSettings settings = AiSettingsStore.Load();
                if (settings == null && internalMode) settings = internalAi ?? AiSettings.Default();
                if (settings == null) {
                    using (var dialog = new AiSettingsDialog()) {
                        if (dialog.ShowDialog() != DialogResult.OK) return;
                        settings = dialog.Selected;
                    }
                }
                Application.Run(new Popup(aiSettings: settings, startInTray: startupLaunch, internalMode: internalMode));
            }
        }
    }
}



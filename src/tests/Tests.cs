using System;
using System.Net;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CorrecteurMail {
    internal sealed class FakeRunner : ICorrectionRunner {
        internal TaskCompletionSource<string> Pending = new TaskCompletionSource<string>();
        internal int Calls, Stops;
        internal bool Fluent;
        internal string LastAdvice, LastQuestion;
        public string Progress { get { return "Réponse simulée"; } }
        public Task<string> Correct(string text, bool fluent, string advice = null) { Calls++; Fluent = fluent; LastAdvice = advice; return Pending.Task; }
        public Task<string> Advise(string text, string question) { Calls++; LastQuestion = question; return Pending.Task; }
        public void Stop() { Stops++; } // Deliberately deliver late answers despite cancellation.
        public void Dispose() { Stop(); }
        internal void Next() { Pending = new TaskCompletionSource<string>(); }
    }
    public sealed class FakeRange { public string Text { get; set; } }
    public sealed class FakeUndoRecord {
        public int Starts, Ends, WritesInside;
        public bool Open;
        public void StartCustomRecord(string name) { Starts++; Open = true; }
        public void EndCustomRecord() { Ends++; Open = false; }
    }
    public sealed class FakeWordApplication { public readonly FakeUndoRecord UndoRecord = new FakeUndoRecord(); }
    public sealed class FakeWordDocument {
        public string Text;
        public readonly FakeWordApplication Application = new FakeWordApplication();
        public readonly List<string> Replaced = new List<string>();
        public FakeWordSpan Range(int start, int end) { return new FakeWordSpan(this, start, end); }
    }
    public sealed class FakeWordSpan {
        private readonly FakeWordDocument document;
        public int Start { get; private set; }
        public int End { get; private set; }
        public FakeWordSpan(FakeWordDocument value, int start, int end) { document = value; Start = start; End = end; }
        public string Text { get { return document.Text.Substring(Start, End - Start); } set {
            if (document.Application.UndoRecord.Open) document.Application.UndoRecord.WritesInside++;
            document.Replaced.Add(Text); document.Text =document.Text.Substring(0, Start) + value + document.Text.Substring(End);
        } }
        // Like Word: a link is \x13 code \x14 label \x15, and Range.Text of the draft shows only the label.
        public FakeFields Fields { get {
            var fields = new FakeFields();
            for (int begin = document.Text.IndexOf('\x13', Start); begin >= 0 && begin < End; begin = document.Text.IndexOf('\x13', begin + 1)) {
                int separator = document.Text.IndexOf('\x14', begin), close = document.Text.IndexOf('\x15', begin);
                fields.Items.Add(new FakeField { Code = new FakeWordSpan(document, begin + 1, separator), Result = new FakeWordSpan(document, separator + 1, close) });
            }
            return fields;
        } }
    }
    public sealed class FakeField { public FakeWordSpan Code { get; set; } public FakeWordSpan Result { get; set; } }
    public sealed class FakeFields {
        public readonly List<FakeField> Items = new List<FakeField>();
        public int Count { get { return Items.Count; } }
        public FakeField Item(int index) { return Items[index - 1]; }
    }
    internal sealed class FakeStartupStore : IStartupStore {
        internal string Value;
        internal int Writes, Removes;
        internal bool DenyWrite;
        public string Read() { return Value; }
        public void Write(string command) { if (DenyWrite) throw new UnauthorizedAccessException(); Value = command; Writes++; }
        public void Remove() { Value = null; Removes++; }
    }
    internal sealed class Rig : IDisposable {
        internal readonly FakeRunner Runner = new FakeRunner();
        internal readonly Popup Window;
        internal string Key = "mail-A", Text = "Bonjour, les document sont prêt.\r";
        internal bool Exists = true, Foreground = true, Menu;
        internal int Reads;
        internal readonly FakeRange Range = new FakeRange();
        internal float DpiScale { get { using (var graphics = Window.CreateGraphics()) return graphics.DpiX / 96f; } }
        internal Size ClientSize {
            get { return new Size((int)Math.Round(Window.ClientSize.Width / DpiScale), (int)Math.Round(Window.ClientSize.Height / DpiScale)); }
            set { Window.ClientSize = new Size((int)Math.Round(value.Width * DpiScale), (int)Math.Round(value.Height * DpiScale)); }
        }
        // liveRange: text written by Appliquer is read back, like Outlook; Range.Text = null stands for Ctrl+Z.
        internal Rig(bool liveRange = false) {
            Window = new Popup(true, Runner, () => { Reads++; return Exists ? new Draft { Key = Key, Text = liveRange && Range.Text != null ? Range.Text : Text, Range = Range } : null; },() => Foreground, appearancePath: Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TEMP", "THEME", "rig.theme.txt"), popupReader: () => Menu);
            Window.Opacity = 0; Window.ShowInTaskbar = false; Window.Show();
        }
        internal T Get<T>(string name) { return (T)typeof(Popup).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Window); }
        internal object Call(string name, params object[] args) { return typeof(Popup).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, args); }
        internal Task Request() { return (Task)Call("Request", false); }
        internal Task Tick() { return (Task)Call("Tick"); }
        internal void Observe() { using (var d = new Draft { Key = Key, Text = Text }) Call("Observe", d); }
        // Ages the last change, as after a long pause: no tick may call the AI on its own.
        internal async Task Wait() { Get<Debounce>("state").Changed = DateTime.UtcNow.AddSeconds(-5); await Tick(); }
        internal async Task Ready(string corrected ="Bonjour, les documents sont prêts.\r") { var task = Request(); Runner.Pending.SetResult(corrected); await task; }
        internal bool Valid { get { return Get<long>("resultVersion") >= 0 && Get<Button>("copy").Enabled; } }
        public void Dispose() { Window.ExitApplication(); Window.Dispose(); }
    }
    internal static class TestProgram {
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool IsProcessDPIAware();
        private static float DisplayScale(Control control) { using (var graphics = control.CreateGraphics()) return graphics.DpiY / 96f; }
        private static readonly List<string> Results = new List<string>();
        private static int failures;
        private static void Check(string name, bool ok) { Results.Add((ok ? "OK   " : "FAIL ") + name); if (!ok) failures++; }
        private static async Task Test(string name, Func<Task<bool>> test) { try { var pending = test(); if (await Task.WhenAny(pending, Task.Delay(5000)) != pending) throw new TimeoutException(name); Check(name, await pending); } catch (Exception ex) { Check(name + " [" + ex.GetBaseException().GetType().Name + ": " + ex.GetBaseException().Message + "]", false); } }
        private static Task<bool> Value(bool value) { return Task.FromResult(value); }
        private static string Event(object value) { return new JavaScriptSerializer().Serialize(new { type = "item.completed", item = new { type = "agent_message", text = new JavaScriptSerializer().Serialize(new { texte = value }) } }); }
        private static string Parse(string value) { return AccountRunner.ReadCorrection(value, new JavaScriptSerializer()); }
        private static AccountRunner ProcessRunner(int timeout = 3000, bool missing = false, AiSettings settings = null, bool connected = true) {
            return new AccountRunner(() => missing ? null : Assembly.GetExecutingAssembly().Location, Path.GetDirectoryName(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)), timeout, settings, () => new AiConnection { Connected = connected, Message = "Compte déconnecté simulé" });
        }
        private static async Task<bool> Error(Task<string> task, string text) { try { await task; return false; } catch (InvalidOperationException ex) { return ex.Message.Contains(text); } }
        private static void FakeCli(bool claude, string[] arguments) {
            // Like an outdated client rejecting an option: quit without reading the request.
            if (Environment.GetEnvironmentVariable("KORECTO_FAKE_EARLY_EXIT") == "1") { Console.Error.WriteLine("error: unknown option '--safe-mode'"); Environment.ExitCode = 2; return; }
            Console.InputEncoding = new UTF8Encoding(false); Console.OutputEncoding = new UTF8Encoding(false);
            string payload = Console.In.ReadToEnd();
            if (payload.Contains("FAKE_AUTH")) { Console.Error.WriteLine("authentication failed"); Environment.ExitCode = 2; return; }
            if (payload.Contains("FAKE_QUOTA")) { Console.Error.WriteLine("rate_limit"); Environment.ExitCode = 2; return; }
            if (payload.Contains("FAKE_EMPTY")) return;
            if (payload.Contains("FAKE_HANG")) { Thread.Sleep(10000); return; }
            Console.WriteLine("{\"type\":42}"); Console.WriteLine("{invalid json");
            string correction = payload.Contains("FAKE_ARGUMENTS") ? String.Join("|", arguments) : payload.Contains("FAKE_PAYLOAD") ? payload.Substring(payload.LastIndexOf("Données utilisateur")) : "Réponse fictive, prête 😊.";
            if (claude && payload.Contains("FAKE_SCHEMA")) {
                int schemaIndex = Array.IndexOf(arguments, "--json-schema");
                new JavaScriptSerializer().DeserializeObject(arguments[schemaIndex + 1]);
                correction = arguments[Array.IndexOf(arguments, "--tools") + 1] == "" ? "Schéma et outils vérifiés" : "ÉCHEC";
            }
            Console.WriteLine(claude ? new JavaScriptSerializer().Serialize(new { type = "result", subtype = "success", is_error = false, structured_output = new { texte = correction } }) : Event(correction)); Console.Out.Flush();
            if (payload.Contains("FAKE_EARLY")) Thread.Sleep(10000);
        }
        [STAThread] private static void Main(string[] args) {
            // Tests never touch the real %LOCALAPPDATA%\Korecto of the user.
            Environment.SetEnvironmentVariable("KORECTO_DONNEES", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TEMP", "DONNEES"));
            if (args.Length > 0 && (args[0] == "exec" || args[0] == "-p")) { FakeCli(args[0] == "-p", args); return; }
            if (args.Length > 0 && args[0] == "login") { Console.Error.WriteLine("Logged in using ChatGPT"); return; }
            if (Array.IndexOf(args, "auth") >= 0) { Console.WriteLine("{\"loggedIn\":true,\"authMethod\":\"claude.ai\"}"); return; }
            if (args.Length > 0 && args[0] == "--check-local-ai") {
                var gpt = LocalAi.Check("chatgpt"); var claude = LocalAi.Check("claude");
                File.WriteAllText(args[1], "ChatGPT : " + gpt.Connected + " — " + gpt.Message + "\r\nClaude : " + claude.Connected + " — " + claude.Message + "\r\nAucun texte envoyé, aucun jeton d’authentification lu ou enregistré par l’application.", Encoding.UTF8); return;
            }
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            using (var host = new Form { Opacity = 0, ShowInTaskbar = false }) {
                host.Load += async (s, e) => {
                    try { await Run(); if (args.Length > 1 && args[1] == "--capture") await Capture(Path.GetDirectoryName(args[0])); if (args.Length > 1 && args[1] == "--compact") { using (var r = new Rig()) { await r.Ready(); r.Window.PerformLayout(); using (var bitmap = new Bitmap(r.Window.Width, r.Window.Height)) { r.Window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(Path.GetDirectoryName(args[0]), "INTERFACE_COMPACTE.png")); } } } }
                    catch (Exception ex) { Check("Erreur globale : " + ex, false); }
                    finally {
                        string target = args.Length > 0 ? args[0] : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RESULTATS.txt");
                        File.WriteAllText(target, "Tests locaux sur données fictives — " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "\r\n" + (Results.Count - failures) + "/" + Results.Count + " réussis ; " + failures + " échec(s).\r\nAucun appel réseau, aucun vrai mail lu ou modifié, presse-papiers intact.\r\n\r\n" + String.Join("\r\n", Results), Encoding.UTF8);
                        Environment.ExitCode = failures == 0 ? 0 : 1; host.Close();
                    }
                }; Application.Run(host);
            }
        }
        private static async Task Run() {
            await Test("Sélection externe : ouverture manuelle sans requête IA ni modification du mail", async () => { using (var r = new Rig()) {
                r.Window.Hide(); r.Window.ImportSelection("Bonjour\v\vTexte choisi 😊"); await r.Wait();
                return r.Window.Visible && r.Get<bool>("manualMode") && r.Get<TextBox>("manualInput").Text == "Bonjour\r\n\r\nTexte choisi 😊" && r.Runner.Calls == 0 && r.Range.Text == null && r.Get<Button>("retry").Enabled && r.Get<Button>("reformulate").Enabled;
            } });
            await Test("Nouvelle sélection : réponse de la correction précédente rejetée", async () => { using (var r = new Rig()) {
                var pending=r.Request(); r.Window.ImportSelection("Nouvelle sélection"); r.Runner.Pending.SetResult("Ancienne correction"); await pending;
                return r.Get<TextBox>("manualInput").Text == "Nouvelle sélection" && !r.Valid && r.Get<TextBox>("result").TextLength == 0;
            } });
            await Test("Sélection externe : Reformuler fonctionne directement sur clic", async () => { using (var r = new Rig()) {
                r.Window.ImportSelection("Bonjour, texte choisi."); var pending=(Task)r.Call("Request",true); r.Runner.Pending.SetResult("Bonjour, voici le texte retenu."); await pending;
                return r.Runner.Calls == 1 && r.Runner.Fluent && r.Valid;
            } });
            await Test("Sélections vides, trop longues ou nulles refusées sans changer le texte", () => { using (var r = new Rig()) {
                r.Window.ImportSelection("Conserver"); foreach (string value in new [] { null, " ", "a\0b", new string('x',30001) }) r.Window.ImportSelection(value);
                return Value(r.Get<TextBox>("manualInput").Text == "Conserver" && r.Runner.Calls == 0 && SelectionAccess.Valid(new string('x',30000)));
            } });
            await Test("Version interne : bouton simulé sans téléchargement ni changement de démarrage Windows", () => {
                using(var window=new Popup(true, internalMode:true)) { window.ShowInTaskbar=false; window.Opacity=0; window.Show(); var button=(Button)typeof(Popup).GetField("updateAvailable",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(window); button.PerformClick(); bool valid=window.Text.Contains("TEST INTERNE") && button.Visible; window.ExitApplication(); return Value(valid); }
            });
            await Test("Mise à jour : bouton discret absent au lancement, visible seulement pour une version récente", async () => { using (var r = new Rig()) {
                var button = r.Get<Button>("updateAvailable"); bool initiallyHidden = !button.Visible;
                await (Task)r.Call("BackgroundUpdateCheck", new Func<string, string>(url => ReleaseJson("v9.0.0", Updater.DownloadPrefix + "v9.0.0/Korecto.exe", 2048, null)));
                var screen = button.RectangleToScreen(button.ClientRectangle); var settings = r.Get<Button>("ai").RectangleToScreen(r.Get<Button>("ai").ClientRectangle);
                bool offered = button.Visible && button.Enabled && button.Text == "Mise à jour disponible" && screen.Bottom <= settings.Top && r.Window.ClientRectangle.Contains(r.Window.RectangleToClient(screen)) && r.Runner.Calls == 0;
                await (Task)r.Call("BackgroundUpdateCheck", new Func<string, string>(url => ReleaseJson("v" + AppBrand.Version, Updater.DownloadPrefix + "v" + AppBrand.Version + "/Korecto.exe", 2048, null)));
                return initiallyHidden && offered && !button.Visible;
            } });
            await Test("Mise à jour : erreur réseau silencieuse sans bouton ni ouverture de fenêtre", async () => { using (var r = new Rig()) {
                r.Window.Hide();
                await (Task)r.Call("BackgroundUpdateCheck", new Func<string, string>(url => { throw new WebException("hors ligne simulé"); }));
                return !r.Window.Visible && !r.Get<Button>("updateAvailable").Visible && r.Runner.Calls == 0;
            } });
            Check("Processus déclaré DPI-aware : pas de grossissement bitmap à l'ouverture", IsProcessDPIAware());
            await Test("Sans mail : fenêtre réduite, Manuel accessible, aucun panneau d'accueil", () => { using (var r = new Rig()) return Value(r.ClientSize.Height <= 240 && !r.Get<TabControl>("textTabs").Visible && r.Get<Button>("manual").Enabled && !r.Get<Button>("expand").Enabled); });
            await Test("Brouillon quitté ou envoyé : Korecto se masque, notification et Manuel conservés", async () => { using (var r = new Rig()) {
                await r.Tick(); await r.Ready(); r.Exists = false; await r.Tick();
                bool menuManual = false; foreach (ToolStripItem item in r.Get<NotifyIcon>("tray").ContextMenuStrip.Items) if (item.Text == "Mode manuel") menuManual = true;
                return !r.Window.Visible && !r.Get<TabControl>("textTabs").Visible && r.Get<long>("resultVersion") < 0 && menuManual && !r.Window.IsDisposed;
            } });
            await Test("Menu ou boîte de dialogue Outlook ouvert : Korecto reste affiché avec sa correction", async () => { using (var r = new Rig()) {
                await r.Tick(); await r.Ready(); r.Exists = false; r.Menu = true; await r.Tick();
                return r.Window.Visible && r.Get<long>("resultVersion") >= 0 && r.Get<TextBox>("result").TextLength > 0;
            } });
            await Test("Nouveau brouillon après masquage : Korecto réapparaît seul", async () => { using (var r = new Rig()) {
                await r.Tick(); await r.Ready(); r.Exists = false; await r.Tick(); bool hidden = !r.Window.Visible;
                r.Exists = true; r.Key = "mail-B"; r.Text = "Nouveau brouillon à corriger.\r"; await r.Tick();
                return hidden && r.Window.Visible && r.Get<TextBox>("manualInput").Text.StartsWith("Nouveau brouillon");
            } });
            await Test("Lancement sans brouillon : fenêtre visible, pas de masquage", async () => { using (var r = new Rig()) { r.Exists = false; await r.Tick(); await r.Tick(); return r.Window.Visible && r.Get<Button>("manual").Enabled; } });
            await Test("Mode Manuel : pas de masquage sans brouillon Outlook", async () => { using (var r = new Rig()) { await r.Ready(); r.Call("SwitchMode"); r.Exists = false; await r.Tick(); return r.Window.Visible; } });
            await Test("Manuel vide : éditeur ouvert et quatre onglets permanents", () => { using (var r = new Rig()) { r.Call("SwitchMode"); var tabs = r.Get<TabControl>("textTabs"); var input = r.Get<TextBox>("manualInput"); input.AppendText("Texte tapé directement"); return Value(!input.ReadOnly && input.Text == "Texte tapé directement" && tabs.Visible && tabs.TabPages.Count == 4 && tabs.SelectedTab == r.Get<TabPage>("inputPage") && r.Runner.Calls == 0); } });
            await Test("Original, corrigé et reformulé sont conservés et copiés séparément", async () => { using (var r = new Rig()) {
                const string correction = "Bonjour, les documents sont prêts.", fluent = "Bonjour, les documents sont prêts à vous être transmis.";
                await r.Ready(correction); string original = r.Get<TextBox>("manualInput").Text;
                r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(fluent); await task;
                var tabs = r.Get<TabControl>("textTabs"); string copied = null;
                r.Call("CopyResult", new Action<string>(value => copied = value)); if (copied != fluent) return false;
                tabs.SelectedTab = r.Get<TabPage>("resultPage"); r.Call("CopyResult", new Action<string>(value => copied = value)); if (copied != correction) return false;
                tabs.SelectedTab = r.Get<TabPage>("inputPage");
                return !r.Get<Button>("copy").Enabled && !r.Get<Button>("apply").Enabled && r.Get<TextBox>("manualInput").Text == original && original.Contains("document sont prêt") && r.Get<TextBox>("manualInput").ReadOnly && r.Get<TextBox>("result").Text == correction && r.Get<TextBox>("reformulated").Text == fluent && r.Runner.Calls == 2;
            } });
            await Test("Appliquer utilise l'onglet choisi et conserve l'original", async () => { using (var r = new Rig()) {
                const string correction = "Les documents sont prêts.", fluent = "Les documents sont disponibles pour validation.";
                await r.Ready(correction); string original = r.Get<TextBox>("manualInput").Text;
                r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(fluent); await task;
                r.Call("Apply"); if (r.Range.Text != fluent + "\r") return false;
                r.Get<TabControl>("textTabs").SelectedTab = r.Get<TabPage>("resultPage"); r.Call("Apply"); await r.Tick();
                return r.Range.Text == correction + "\r" && r.Get<TextBox>("manualInput").Text == original && r.Get<TextBox>("reformulated").Text == fluent;
            } });
            await Test("Modifier l'original manuel garde Corrigé et Reformulé, Corriger réactivé", async () => { using (var r = new Rig()) {
                r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Les document sont prêt."; await r.Ready("Les documents sont prêts.");
                r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Les documents sont disponibles."); await task;
                r.Get<TextBox>("manualInput").AppendText(" Merci.");
                return r.Get<TextBox>("result").Text == "Les documents sont prêts." && r.Get<TextBox>("reformulated").Text == "Les documents sont disponibles." && !r.Valid && r.Get<Button>("retry").Enabled;
            } });
            await Test("Une reformulation tardive ne remplit pas les onglets d'un autre mail", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Key = "mail-B"; r.Text = "Texte du mail B"; await r.Tick(); r.Runner.Pending.SetResult("Ancienne reformulation de A"); await task;
                return r.Get<TextBox>("manualInput").Text == r.Text && r.Get<TextBox>("result").TextLength == 0 && r.Get<TextBox>("reformulated").TextLength == 0 && !r.Valid;
            } });
            await Test("La reformulation ne détruit pas une variante précédente lorsqu'elle répète la source", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, les documents sont disponibles."); await task; string previous = r.Get<TextBox>("reformulated").Text;
                r.Runner.Next(); task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(r.Text); await task;
                return r.Get<TextBox>("reformulated").Text == previous && r.Get<TextBox>("manualInput").Text.Contains("document sont prêt") && r.Get<Label>("status").Text.Contains("Aucune reformulation proposée");
            } });
            var now = DateTime.UtcNow; var d = new Debounce();
            Check("Premier brouillon détecté", d.Observe("A", "Texte", now));
            long v = d.Version;
            Check("Lecture identique conserve la version du texte", !d.Observe("A", "Texte", now.AddSeconds(15)) && d.Version == v);
            Check("Deux mails identiques restent distincts", d.Observe("B", "Texte", now) && !d.Matches(v, "B", "Texte"));
            d.Observe("B", "\r\n \t", now); Check("Brouillon blanc reconnu", String.IsNullOrWhiteSpace(d.Text));
            d.Observe("A", "Texte", now); v = d.Version; d.Observe("B", "Texte", now); d.Observe("A", "Texte", now);
            Check("Aller-retour A/B/A invalide la vieille réponse", !d.Matches(v, "A", "Texte"));
            await Test("Frappe visible en direct sans aucune requête IA", async () => { using (var r = new Rig()) { for (int i = 0; i < 20; i++) { r.Text = "Bonjour " + new string('a', i + 1); await r.Tick(); if (r.Get<TextBox>("manualInput").Text != r.Text || r.Runner.Calls != 0) return false; } return r.Get<Button>("retry").Enabled && !r.Valid; } });
            string[] numbers = { "1 250,50 €", "26/10/2026 à 14h30", "Réf. A-42", "3,5 m²", "-100 €", "10 %", "0,001", "☎ 06 12 34 56 78" };
            foreach (string n in numbers) Check("Chiffres conservés : " + n, Integrity.SameNumbers(n, "Bonjour " + n + "."));
            Check("Montant modifié refusé", !Integrity.SameNumbers("1 250,50 €", "1 200,50 €"));
            Check("Date modifiée refusée", !Integrity.SameNumbers("26/10/2026", "27/10/2026"));
            Check("Chiffre ajouté refusé", !Integrity.SameNumbers("Bonjour", "Bonjour 10"));
            Check("Chiffre supprimé refusé", !Integrity.SameNumbers("10 jours", "des jours"));
            Check("Signe négatif supprimé refusé", !Integrity.SameNumbers("-100 €", "100 €"));
            Check("Signe négatif espacé supprimé refusé", !Integrity.SameNumbers("- 100 €", "100 €"));
            Check("Signe moins Unicode supprimé refusé", !Integrity.SameNumbers("−100 €", "100 €"));
            Check("Signe plus supprimé refusé", !Integrity.SameNumbers("+100 €", "100 €"));
            Check("Ordre des chiffres modifié refusé", !Integrity.SameNumbers("12 puis 34", "34 puis 12"));
            await Test("JSON correct accepté", () => Value(Parse(Event("Été, prêt 😊")) == "Été, prêt 😊"));
            await Test("Autre événement ignoré", () => Value(Parse("{\"type\":\"turn.started\"}") == null));
            await Test("Réponse vide ignorée", () => Value(Parse(Event(" \r\n")) == null));
            await Test("Réponse numérique ignorée", () => Value(Parse(Event(42)) == null));
            await Test("Type événement incorrect ignoré sans exception", () => Value(Parse("{\"type\":42}") == null));
            await Test("Type item incorrect ignoré sans exception", () => Value(Parse("{\"type\":\"item.completed\",\"item\":{\"type\":42}}") == null));
            await Test("Texte item incorrect ignoré sans exception", () => Value(Parse("{\"type\":\"item.completed\",\"item\":{\"type\":\"agent_message\",\"text\":42}}") == null));
            await Test("JSON tronqué ignoré sans exception", () => Value(Parse("{\"type\":") == null));
            await Test("Aperçu automatique et saisie manuelle", () => { using (var r = new Rig()) return Value(r.Window.TestDraftPreview()); });
            await Test("Saisie Outlook : aperçu toujours visible, sans écran vide ni déplacement", async () => { using (var r = new Rig()) {
                await r.Tick(); var tabs = r.Get<TabControl>("textTabs"); var preview = r.Get<TextBox>("manualInput");
                int visibility = 0, empty = 0, moves = 0, sizes = 0, updates = 0;
                tabs.VisibleChanged += (s,e) => visibility++;
                preview.TextChanged += (s,e) => { updates++; if (preview.Text.Length == 0) empty++; };
                r.Window.LocationChanged += (s,e) => moves++; r.Window.SizeChanged += (s,e) => sizes++;
                for (int i = 0; i < 30; i++) { r.Text += "a"; await r.Tick(); }
                return visibility == 0 && empty == 0 && moves == 0 && sizes == 0 && updates == 30 && tabs.Visible && r.Runner.Calls == 0;
            } });
            await Test("Saisie après correction : Corrigé reste affiché, Appliquer bloqué, aperçu à jour", async () => { using (var r = new Rig()) {
                await r.Ready(); var tabs = r.Get<TabControl>("textTabs"); var preview = r.Get<TextBox>("manualInput"); int visibility = 0, empty = 0;
                tabs.VisibleChanged += (s,e) => visibility++; preview.TextChanged += (s,e) => { if (preview.Text.Length == 0) empty++; };
                r.Text += " Ajout"; await r.Tick();
                return visibility == 0 && empty == 0 && preview.Text == r.Text.Replace("\r", "\r\n") && tabs.SelectedTab == r.Get<TabPage>("resultPage") && r.Get<TextBox>("result").Text.StartsWith("Bonjour, les documents sont prêts.")
                    && !r.Valid && !r.Get<Button>("apply").Enabled && r.Get<Button>("copy").Enabled && r.Get<Button>("retry").Enabled && r.Get<Label>("status").Text.Contains("gardée") && r.Runner.Calls == 1;
            } });
            await Test("Proposition gardée : Appliquer n'écrase pas les nouvelles modifications", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Text += " Ajout"; await r.Tick(); r.Call("Apply"); return r.Range.Text == null;
            } });
            await Test("Proposition gardée : Copier donne le texte précédent avec un message clair", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Text += " Ajout"; await r.Tick(); string copied = null;
                r.Call("CopyResult", new Action<string>(value => copied = value));
                return copied == r.Get<TextBox>("result").Text && copied.StartsWith("Bonjour, les documents sont prêts.") && r.Get<Label>("status").Text.Contains("précédente");
            } });
            await Test("Proposition gardée : jamais copiée pour un autre mail", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Text += " Ajout"; await r.Tick(); r.Key = "mail-B"; string copied = null;
                r.Call("CopyResult", new Action<string>(value => copied = value)); return copied == null;
            } });
            await Test("Changer de mail efface Corrigé et Reformulé de l'autre mail", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Key = "mail-B"; r.Text = "Texte du second mail.\r"; await r.Tick();
                return r.Get<TextBox>("result").TextLength == 0 && r.Get<TextBox>("reformulated").TextLength == 0 && !r.Get<Button>("copy").Enabled;
            } });
            await Test("Erreur après modification : proposition précédente conservée, pas remplacée par l'erreur", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Text += " Ajout"; await r.Tick(); r.Runner.Next(); var task = r.Request(); r.Runner.Pending.SetException(new InvalidOperationException("Erreur simulée")); await task;
                return r.Get<TextBox>("result").Text.StartsWith("Bonjour, les documents sont prêts.") && r.Get<Label>("status").Text.Contains("conservée") && r.Get<Button>("retry").Enabled;
            } });
            await Test("Aperçu inchangé : aucun remplacement du texte ni changement de panneau", async () => { using (var r = new Rig()) {
                await r.Tick(); int changes = 0; r.Get<TextBox>("manualInput").TextChanged += (s,e) => changes++; r.Get<TabControl>("textTabs").VisibleChanged += (s,e) => changes++;
                for (int i = 0; i < 20; i++) await r.Tick(); return changes == 0 && r.Runner.Calls == 0;
            } });
            await Test("Saisie manuelle : le panneau et le curseur restent stables", () => { using (var r = new Rig()) {
                r.Call("SwitchMode"); var input = r.Get<TextBox>("manualInput"); var tabs = r.Get<TabControl>("textTabs"); int visibility = 0, moves = 0;
                tabs.VisibleChanged += (s,e) => visibility++; r.Window.LocationChanged += (s,e) => moves++;
                for (int i = 0; i < 30; i++) input.AppendText("a");
                return Value(visibility == 0 && moves == 0 && tabs.SelectedTab == r.Get<TabPage>("inputPage") && input.SelectionStart == input.TextLength && r.Runner.Calls == 0);
            } });
            await Test("Réponse valide visible et actions actives", async () => { using (var r = new Rig()) { await r.Ready(); return r.Valid && r.Get<Button>("apply").Enabled && r.Get<TextBox>("result").Text.Contains("prêts"); } });
            await Test("Texte corrigé : Corriger grisé et appel direct sans doublon", async () => { using (var r = new Rig()) { await r.Ready(); await r.Tick(); bool disabled = !r.Get<Button>("retry").Enabled; await r.Request(); return disabled && r.Runner.Calls == 1 && r.Valid && r.Get<Button>("reformulate").Enabled; } });
            await Test("Éditer le brouillon réactive Corriger", async () => { using (var r = new Rig()) { await r.Ready(); r.Text += " Modification"; await r.Tick(); return r.Get<Button>("retry").Enabled && !r.Valid && r.Runner.Calls == 1; } });
            await Test("Manuel : correction grisée puis réactivée à la frappe", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Bonjour."; await r.Ready("Bonjour."); await r.Tick(); bool disabled = !r.Get<Button>("retry").Enabled; r.Get<TextBox>("manualInput").AppendText(" Merci."); return disabled && r.Get<Button>("retry").Enabled && !r.Valid; } });
            await Test("Après Appliquer, le même texte n'est pas recorrigé", async () => { using (var r = new Rig()) { await r.Ready(); r.Call("Apply"); await r.Tick(); await r.Request(); return !r.Get<Button>("retry").Enabled && r.Runner.Calls == 1; } });
            await Test("Reformulation identique à la source : message explicite", async () => { using (var r = new Rig()) { var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(r.Text); await task; return r.Get<Label>("status").Text.Contains("Aucune reformulation proposée") && r.Valid && !r.Get<Button>("retry").Enabled; } });
            await Test("Reformulation identique à la correction : pas de faux changement", async () => { using (var r = new Rig()) { await r.Ready(); string previous = r.Get<TextBox>("result").Text; r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(previous); await task; await r.Tick(); return r.Get<Label>("status").Text.Contains("Aucune reformulation proposée") && r.Get<TextBox>("result").Text == previous && r.Valid && !r.Get<Button>("retry").Enabled; } });
            await Test("Reformulation ne variant que les espaces : message explicite", async () => { using (var r = new Rig()) { r.Text = "Bonjour, Carla.\rMerci."; var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(" Bonjour,  Carla.\n\nMerci. "); await task; return r.Get<Label>("status").Text.Contains("Aucune reformulation proposée") && r.Valid; } });
            await Test("Reformulation différente : onglet et proposition distincts", async () => { using (var r = new Rig()) { var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, je vous transmets les documents prêts."); await task; return r.Get<TabControl>("textTabs").SelectedTab == r.Get<TabPage>("reformulationPage") && r.Get<Label>("status").Text.StartsWith("Reformulation proposée") && r.Valid; } });
            await Test("Reformulation en erreur : correction précédente conservée", async () => { using (var r = new Rig()) { await r.Ready(); string previous = r.Get<TextBox>("result").Text; r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetException(new InvalidOperationException("Quota simulé")); await task; return r.Get<TextBox>("result").Text == previous && r.Valid && r.Get<Label>("status").Text.Contains("précédente conservée") && !r.Get<Button>("retry").Enabled; } });
            await Test("Reformulation qui répète la source préserve la correction utile", async () => { using (var r = new Rig()) { await r.Ready(); string previous = r.Get<TextBox>("result").Text; r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult(r.Text); await task; return r.Get<TextBox>("result").Text == previous && r.Valid && r.Get<Label>("status").Text.Contains("Aucune reformulation proposée"); } });
            await Test("Proposition aux chiffres modifiés : Corriger reste disponible", async () => { using (var r = new Rig()) { r.Text = "Montant 100 €"; await r.Ready("Montant 200 €"); return r.Get<Button>("retry").Enabled && !r.Get<Button>("apply").Enabled; } });
            await Test("Nom Korecto dans la fenêtre, la notification et les paramètres", () => { using (var r = new Rig()) using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => null, new StartupPreference(new FakeStartupStore()))) return Value(r.Window.Text.StartsWith("Korecto ·") && r.Get<NotifyIcon>("tray").Text == "Korecto" && dialog.Text == "Paramètres · Korecto"); });
            await Test("Double clic ne lance qu'une requête", async () => { using (var r = new Rig()) { var t = r.Request(); await r.Request(); r.Runner.Pending.SetResult("Bonjour."); await t; return r.Runner.Calls == 1; } });
            await Test("Longue pause dans Outlook : aucune correction sans clic", async () => { using (var r = new Rig()) { await r.Tick(); r.Get<Debounce>("state").Changed = DateTime.UtcNow.AddMinutes(-10); for (int i = 0; i < 5; i++) await r.Tick(); return r.Runner.Calls == 0 && r.Get<Button>("retry").Enabled && r.Get<Label>("status").Text.Contains("cliquez sur Corriger"); } });
            await Test("Corriger lance la correction tout de suite, une seule fois", async () => { using (var r = new Rig()) { await r.Tick(); var task = r.Request(); bool started = r.Runner.Calls == 1; r.Runner.Pending.SetResult("Bonjour, les documents sont prêts.\r"); await task; for (int i = 0; i < 4; i++) await r.Tick(); return started && r.Runner.Calls == 1 && r.Valid; } });
            await Test("Édition après un clic : réponse périmée rejetée, nouveau clic nécessaire", async () => { using (var r = new Rig()) { await r.Tick(); var first = r.Request(); r.Text = "Nouveau texte"; await r.Tick(); r.Runner.Pending.SetResult("Ancienne réponse"); await first; await r.Wait(); if (r.Valid || r.Runner.Calls != 1) return false; r.Runner.Next(); var second = r.Request(); r.Runner.Pending.SetResult("Nouveau texte corrigé"); await second; return r.Valid && r.Runner.Calls == 2 && r.Get<TextBox>("result").Text == "Nouveau texte corrigé"; } });
            await Test("Mode manuel : même après une longue pause, aucun appel automatique", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Texte manuel"; r.Get<Debounce>("state").Changed = DateTime.UtcNow.AddMinutes(-10); await r.Tick(); return r.Runner.Calls == 0 && r.Get<Button>("retry").Enabled; } });
            await Test("Après une erreur, pas de répétition sans clic", async () => { using (var r = new Rig()) { await r.Tick(); var task = r.Request(); r.Runner.Pending.SetException(new InvalidOperationException("Erreur simulée")); await task; for (int i = 0; i < 3; i++) await r.Wait(); return r.Runner.Calls == 1 && !r.Valid && r.Get<Label>("status").Text.Contains("Corriger"); } });
            await Test("Changer de mail met à jour l’aperçu sans appel IA", async () => { using (var r = new Rig()) { await r.Tick(); r.Key = "mail-B"; r.Text = "Texte du second mail"; await r.Tick(); return r.Runner.Calls == 0 && r.Get<TextBox>("manualInput").Text == r.Text; } });
            await Test("Pause puis reprise restaurent l’aperçu sans appel IA", async () => { using (var r = new Rig()) { await r.Tick(); r.Call("TogglePause"); r.Call("TogglePause"); await r.Tick(); return r.Runner.Calls == 0 && r.Get<TextBox>("manualInput").Text == "Bonjour, les document sont prêt.\r\n"; } });
            await Test("Bascule A vers B pendant requête", async () => { using (var r = new Rig()) { var t = r.Request(); r.Key = "mail-B"; r.Observe(); r.Runner.Pending.SetResult("Réponse de A"); await t; return !r.Valid && r.Runner.Stops > 0; } });
            await Test("Réponse périmée préserve l'aperçu du nouveau mail", async () => { using (var r = new Rig()) { var t = r.Request(); r.Key = "mail-B"; r.Text = "Nouveau mail"; r.Observe(); r.Runner.Pending.SetResult("Réponse de A"); await t; return !r.Valid && r.Get<TextBox>("manualInput").Text == "Nouveau mail"; } });
            await Test("Bascule A vers B sans attendre le timer", async () => { using (var r = new Rig()) { var t = r.Request(); r.Key = "mail-B"; r.Runner.Pending.SetResult("Réponse de A"); await t; return !r.Valid; } });
            await Test("A/B/A pendant requête même texte", async () => { using (var r = new Rig()) { var t = r.Request(); r.Key = "mail-B"; r.Observe(); r.Key = "mail-A"; r.Observe(); r.Runner.Pending.SetResult("Réponse périmée"); await t; return !r.Valid; } });
            await Test("Brouillon fermé pendant requête", async () => { using (var r = new Rig()) { var t = r.Request(); r.Exists = false; r.Runner.Pending.SetResult("Réponse périmée"); await t; return !r.Valid; } });
            await Test("Texte modifié juste avant réception", async () => { using (var r = new Rig()) { var t = r.Request(); r.Text += "Modification"; r.Runner.Pending.SetResult("Réponse périmée"); await t; return !r.Valid; } });
            await Test("Copier revalide le mail actif immédiatement", async () => { using (var r = new Rig()) { await r.Ready(); r.Key = "mail-B"; string copied = null; r.Call("CopyResult", new Action<string>(x => copied = x)); return copied == null && !r.Valid; } });
            await Test("Copier refuse un brouillon fermé", async () => { using (var r = new Rig()) { await r.Ready(); r.Exists = false; string copied = null; r.Call("CopyResult", new Action<string>(x => copied = x)); return copied == null; } });
            await Test("Appliquer refuse un autre mail", async () => { using (var r = new Rig()) { await r.Ready(); r.Key = "mail-B"; r.Call("Apply"); return r.Range.Text == null && !r.Valid; } });
            await Test("Appliquer refuse une modification hors timer", async () => { using (var r = new Rig()) { await r.Ready(); r.Text += "Merci"; r.Call("Apply"); return r.Range.Text == null; } });
            await Test("Appliquer écrit seulement la plage du brouillon", async () => { using (var r = new Rig()) { await r.Ready(); r.Call("Apply"); return r.Range.Text == "Bonjour, les documents sont prêts.\r"; } });
            await Test("Appliquer conserve les paragraphes et exclut la signature", () => {
                string original = "Bonjour,\rLes document sont prêt.\rMerci.\r";
                var document = new FakeWordDocument { Text = original + "SIGNATURE\r" };
                var draft = new Draft { Text = original, Document = document, Range = document.Range(0, original.Length) };
                Popup.ReplaceDraftText(draft, "Bonjour,\rLes documents sont prêts.\rMerci.\r");
                return Value(document.Text == "Bonjour,\rLes documents sont prêts.\rMerci.\rSIGNATURE\r" && document.Replaced.Count == 1 && !document.Replaced[0].Contains("\r"));
            });
            await Test("Appliquer une reformulation avec nouveaux paragraphes garde la marque finale", () => {
                string original = "Bonjour, texte.\r"; var document = new FakeWordDocument { Text = original };
                Popup.ReplaceDraftText(new Draft { Text = original, Document = document, Range = document.Range(0, original.Length) }, "Bonjour,\rNouveau texte.\r");
                return Value(document.Text == "Bonjour,\rNouveau texte.\r" && document.Replaced.Count == 1 && !document.Replaced[0].EndsWith("\r"));
            });
            await Test("Appliquer un texte identique ne touche pas au document", () => {
                var document = new FakeWordDocument { Text = "Bonjour 😊\r" };
                Popup.ReplaceDraftText(new Draft { Text = document.Text, Document = document, Range = document.Range(0, document.Text.Length) }, document.Text);
                return Value(document.Replaced.Count == 0);
            });
            await Test("Appliquer corrige autour d'un lien et d'une mention, gardés tels quels même avec une faute", () => {
                string mention = "\x13HYPERLINK \"mailto:n.sacquet@exemple.fr\"\x14@Nicolas SACQUET\x15", link = "\x13HYPERLINK \"https://exemple.fr\"\x14lein\x15";
                var document = new FakeWordDocument { Text = "Bonjour " + mention + " nous a demander le plan.\rVoir le " + link + " et les document.\rSIGNATURE\r" };
                string original = "Bonjour @Nicolas SACQUET nous a demander le plan.\rVoir le lein et les document.\r";
                var draft = new Draft { Text = original, Document = document, Range = document.Range(0, document.Text.Length - "SIGNATURE\r".Length) };
                int kept = Popup.ReplaceDraftText(draft, "Bonjour Nicolas Sacquet nous a demandé le plan.\rVoir le lien et les documents.\r");
                bool untouched = true; foreach (string replaced in document.Replaced) if (replaced.IndexOfAny(new [] { '\x13', '\x14', '\x15' }) >= 0) untouched = false;
                return Value(kept > 0 && untouched && document.Text == "Bonjour " + mention + " nous a demandé le plan.\rVoir le " + link + " et les documents.\rSIGNATURE\r");
            });
            await Test("Appliquer : lien supprimé par l'IA remis en place, texte avant et après corrigé", () => {
                string link = "\x13HYPERLINK \"https://exemple.fr\"\x14le site\x15";
                var document = new FakeWordDocument { Text = "Les plan sont sur " + link + " depuis hier.\r" };
                var draft = new Draft { Text = "Les plan sont sur le site depuis hier.\r", Document = document, Range = document.Range(0, document.Text.Length) };
                Popup.ReplaceDraftText(draft, "Les plans sont en ligne depuis hier.\r");
                return Value(document.Text.Contains(link) && document.Text.StartsWith("Les plans sont"));
            });
            await Test("Appliquer refuse un lien que Word décrit autrement, sans rien modifier", () => {
                var document = new FakeWordDocument { Text = "Voir le \x13HYPERLINK \"https://exemple.fr\"\x14lien\x15.\rLes document sont prêt.\r" };
                string before = document.Text, refusal = null;
                var draft = new Draft { Text = "Voir le site.\rLes document sont prêt.\r", Document = document, Range = document.Range(0, document.Text.Length) };
                try { Popup.ReplaceDraftText(draft, "Voir le site.\rLes documents sont prêts.\r"); } catch (InvalidOperationException ex) { refusal = ex.Message; }
                return Value(refusal == Popup.ComplexDraft && document.Text == before && document.Replaced.Count == 0);
            });
            await Test("Appliquer refuse de toucher une image incorporée, sans écriture partielle", () => {
                string original = "Bonjour,\rPhoto \x01 jointe.\rMerci\r"; var document = new FakeWordDocument { Text = original }; string refusal = null;
                try { Popup.ReplaceDraftText(new Draft { Text = original, Document = document, Range = document.Range(0, original.Length) }, "Bonjour !\rPhoto jointe.\rMerci\r"); } catch (InvalidOperationException ex) { refusal = ex.Message; }
                return Value(refusal == Popup.ComplexDraft && document.Text == original && document.Replaced.Count == 0);
            });
            await Test("Appliquer corrige autour d'une image incorporée laissée intacte", () => {
                string original = "Les document sont prêt.\rPhoto \x01.\r"; var document = new FakeWordDocument { Text = original };
                Popup.ReplaceDraftText(new Draft { Text = original, Document = document, Range = document.Range(0, original.Length) }, "Les documents sont prêts.\rPhoto \x01.\r");
                return Value(document.Text == "Les documents sont prêts.\rPhoto \x01.\r");
            });
            await Test("Appliquer : une seule étape d'annulation (Ctrl+Z) dans Outlook", () => {
                string original = "Bonjour,\rLes document sont prêt.\rMerci de confirmé.\r"; var document = new FakeWordDocument { Text = original };
                Popup.ReplaceDraftText(new Draft { Text = original, Document = document, Range = document.Range(0, original.Length) }, "Bonjour,\rLes documents sont prêts.\rMerci de confirmer.\r");
                var undo = document.Application.UndoRecord;
                return Value(undo.Starts == 1 && undo.Ends == 1 && undo.WritesInside == 2 && !undo.Open);
            });
            await Test("Ctrl+Z puis Ctrl+Y après Appliquer : aucune nouvelle demande à l'IA", async () => { using (var r = new Rig(true)) {
                await r.Tick(); await r.Ready(); r.Call("Apply"); await r.Tick();
                bool applied = r.Range.Text == "Bonjour, les documents sont prêts.\r"; int calls = r.Runner.Calls;
                r.Range.Text = null; await r.Tick(); await r.Wait();
                bool restored = r.Valid && r.Get<TextBox>("result").Text.StartsWith("Bonjour, les documents sont prêts.") && r.Get<Label>("status").Text.Contains("sans nouvelle demande");
                r.Range.Text = "Bonjour, les documents sont prêts.\r"; await r.Tick(); await r.Wait();
                return applied && restored && r.Runner.Calls == calls;
            } });
            await Test("Retour au texte déjà corrigé : correction réaffichée sans appel", async () => { using (var r = new Rig()) {
                await r.Tick(); await r.Ready(); int calls = r.Runner.Calls; string original = r.Text;
                r.Text = original + "Merci"; await r.Tick(); r.Text = original; await r.Tick(); await r.Wait();
                return r.Runner.Calls == calls && r.Valid && !r.Get<Button>("retry").Enabled;
            } });
            await Test("Annuler une reformulation appliquée : aucune demande, Corriger possible", async () => { using (var r = new Rig(true)) {
                await r.Tick(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, je vous transmets les documents prêts."); await task;
                r.Call("Apply"); await r.Tick(); bool applied = r.Range.Text == "Bonjour, je vous transmets les documents prêts.\r"; int calls = r.Runner.Calls;
                r.Range.Text = null; await r.Tick(); await r.Wait();
                return applied && r.Runner.Calls == calls && r.Get<Label>("status").Text.Contains("déjà traité") && r.Get<Button>("retry").Enabled;
            } });
            await Test("Texte jamais vu : aucune demande sans clic, Corriger actif", async () => { using (var r = new Rig()) {
                await r.Tick(); await r.Ready(); int calls = r.Runner.Calls;
                r.Text = "Autre texte à corrigé.\r"; await r.Tick(); await r.Wait();
                return r.Runner.Calls == calls && r.Get<Button>("retry").Enabled;
            } });
            await Test("Reformulation aux chiffres modifiés ne réactive pas Corriger",async () => { using (var r = new Rig()) {
                await r.Ready(); r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, les 2 documents sont prêts."); await task;
                return !r.Get<Button>("retry").Enabled && r.Get<long>("correctedVersion") == r.Get<Debounce>("state").Version && r.Get<Label>("status").Text.Contains("chiffres");
            } });
            await Test("Modification : Reformuler reste disponible, proposition périmée", async () => { using (var r = new Rig()) { await r.Ready(); r.Text += "Merci"; r.Observe(); return r.Get<Button>("reformulate").Enabled && !r.Valid; } });
            await Test("Reformuler disponible sans corriger d'abord, sans demande avant le clic", async () => { using (var r = new Rig()) { await r.Tick(); bool ready = r.Get<Button>("reformulate").Enabled && r.Runner.Calls == 0; var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, les documents sont prêts."); await task; return ready && r.Runner.Calls == 1 && r.Runner.Fluent && r.Valid; } });
            await Test("Pause bloque copie même résultat restauré", async () => { using (var r = new Rig()) { await r.Ready(); r.Call("TogglePause"); return !r.Valid && !r.Get<Button>("retry").Enabled && !r.Get<Button>("reformulate").Enabled; } });
            await Test("Pause/reprise rapide rejette la réponse annulée", async () => { using (var r = new Rig()) { var t = r.Request(); r.Call("TogglePause"); r.Call("TogglePause"); r.Runner.Pending.SetResult("Ancienne réponse"); await t; return !r.Valid; } });
            await Test("Pause interrompt la requête et ignore réponse", async () => { using (var r = new Rig()) { var t = r.Request(); r.Call("TogglePause"); r.Runner.Pending.SetResult("Ancienne réponse"); await t; return !r.Valid && r.Runner.Stops > 0; } });
            await Test("Reprise manuelle réactive Corriger", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Bonjour"; r.Call("TogglePause"); r.Call("TogglePause"); await r.Tick(); return r.Get<Button>("retry").Enabled; } });
            await Test("Mode manuel conserve source et copie correction", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); var input = r.Get<TextBox>("manualInput"); input.Text = "Les document sont prêt."; await r.Ready("Les documents sont prêts."); string copied = null; r.Call("CopyResult", new Action<string>(x => copied = x)); return copied == "Les documents sont prêts." && input.Text == "Les document sont prêt."; } });
            await Test("Édition manuelle pendant requête", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Bonjour"; var t = r.Request(); r.Get<TextBox>("manualInput").Text = "Autre texte"; r.Runner.Pending.SetResult("Ancienne réponse"); await t; return !r.Valid; } });
            await Test("Passage manuel pendant requête Outlook", async () => { using (var r = new Rig()) { var t = r.Request(); r.Call("SwitchMode"); r.Runner.Pending.SetResult("Ancienne réponse"); await t; return !r.Valid && r.Get<bool>("manualMode"); } });
            await Test("Erreur moteur affiche message et permet réessai", async () => { using (var r = new Rig()) { var t = r.Request(); r.Runner.Pending.SetException(new InvalidOperationException("Connexion indisponible simulée")); await t; bool error = !r.Valid && r.Get<TextBox>("result").Text.Contains("Connexion indisponible"); r.Runner.Next(); await r.Ready(); return error && r.Valid && r.Runner.Calls == 2; } });
            await Test("Réponse blanche n'active aucune action", async () => { using (var r = new Rig()) { await r.Ready(" \r\n"); return !r.Valid && !r.Get<Button>("apply").Enabled; } });
            await Test("Réponse nulle traitée comme erreur", async () => { using (var r = new Rig()) { await r.Ready(null); return !r.Valid && r.Get<Button>("retry").Enabled; } });
            await Test("Chiffres différents bloquent Appliquer après timer", async () => { using (var r = new Rig()) { r.Text = "Montant 100 €"; await r.Ready("Montant 200 €"); await r.Tick(); r.Call("Apply"); return !r.Get<Button>("apply").Enabled && r.Range.Text == null; } });
            await Test("Texte blanc ne déclenche pas le moteur", async () => { using (var r = new Rig()) { r.Text = "\r\n \t"; await r.Request(); return r.Runner.Calls == 0; } });
            await Test("30 001 caractères refusés avant transmission", async () => { using (var r = new Rig()) { r.Text = new string('a', 30001); await r.Request(); return r.Runner.Calls == 0 && !r.Get<bool>("busy"); } });
            await Test("30 000 caractères acceptés sans troncature", async () => { using (var r = new Rig()) { r.Text = new string('a', 30000); await r.Ready(r.Text); return r.Runner.Calls == 1 && r.Get<TextBox>("result").Text.Length == 30000; } });
            await Test("Accents, emoji et retours à la ligne", async () => { using (var r = new Rig()) { r.Text = "Été 😊\rDeuxième ligne\nFin\r\n"; await r.Ready(r.Text); return r.Get<TextBox>("result").Text == "Été 😊\r\nDeuxième ligne\r\nFin\r\n"; } });
            await Test("Deux sauts manuels Outlook affichent une ligne vide, sans carrés ni modification du brouillon", async () => { using (var r = new Rig()) {
                string source = "Bonjour M. DUEZ\v\vSuite 😊\r";
                r.Text = source; await r.Tick();
                bool preview = r.Get<TextBox>("manualInput").Text == "Bonjour M. DUEZ\r\n\r\nSuite 😊\r\n" && r.Get<Debounce>("state").Text == source && r.Range.Text == null;
                await r.Ready(source);
                return preview && r.Get<TextBox>("result").Text == "Bonjour M. DUEZ\r\n\r\nSuite 😊\r\n" && r.Valid && r.Text == source;
            } });
            await Test("Séparateurs Unicode affichés en retours à la ligne dans les trois onglets", async () => { using (var r = new Rig()) {
                r.Text = "Bonjour\u2028Suite\u2029Fin\r"; await r.Tick(); await r.Ready(r.Text);
                r.Runner.Next(); var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour\u2028Autre tournure\u2029Fin\r"); await task;
                return r.Get<TextBox>("manualInput").Text == "Bonjour\r\nSuite\r\nFin\r\n" && r.Get<TextBox>("result").Text == "Bonjour\r\nSuite\r\nFin\r\n" && r.Get<TextBox>("reformulated").Text == "Bonjour\r\nAutre tournure\r\nFin\r\n";
            } });
            await Test("Onglets Corrigé et Reformulé inaccessibles avant une première proposition", async () => { using (var r = new Rig()) {
                var tabs = r.Get<TabControl>("textTabs"); var input = r.Get<TabPage>("inputPage"); var handle = tabs.Handle;
                tabs.SelectedTab = r.Get<TabPage>("resultPage"); bool blocked = tabs.SelectedTab == input;
                tabs.SelectedTab = r.Get<TabPage>("reformulationPage"); blocked = blocked && tabs.SelectedTab == input;
                await r.Ready(); tabs.SelectedTab = r.Get<TabPage>("reformulationPage");
                return blocked && tabs.SelectedTab == r.Get<TabPage>("reformulationPage");
            } });
            Check("Nouveautés : la version en cours figure dans la liste", AppAssets.Read(AppDomain.CurrentDomain.BaseDirectory, "NOUVEAUTES.txt").StartsWith("Version " + AppBrand.Version));
            await Test("Onglet Avis : ouvert dès qu'il y a du texte, fermé sans texte, aucune demande avant le clic", async () => {
                using (var empty = new Rig()) {
                    empty.Exists = false; await empty.Tick(); var tabs = empty.Get<TabControl>("textTabs"); var handle = tabs.Handle;
                    tabs.SelectedTab = empty.Get<TabPage>("advicePage"); if (tabs.SelectedTab != empty.Get<TabPage>("inputPage")) return false;
                }
                using (var r = new Rig()) {
                    await r.Tick(); var tabs = r.Get<TabControl>("textTabs"); var handle = tabs.Handle; tabs.SelectedTab = r.Get<TabPage>("advicePage");
                    return tabs.SelectedTab == r.Get<TabPage>("advicePage") && r.Runner.Calls == 0 && r.Get<Button>("adviceAsk").Enabled && !r.Get<Button>("adviceReformulate").Enabled;
                }
            });
            await Test("Demander l'avis : question transmise, réponse affichée, mail intact", async () => { using (var r = new Rig()) {
                await r.Tick(); r.Get<TextBox>("adviceQuestion").Text = "Mon mail est-il trop agressif ?";
                var task = (Task)r.Call("RequestAdvice"); r.Runner.Pending.SetResult("Oui, trop sec : adoucir l'ouverture."); await task;
                return r.Runner.Calls == 1 && r.Runner.LastQuestion == "Mon mail est-il trop agressif ?" && r.Get<TextBox>("adviceBox").Text == "Oui, trop sec : adoucir l'ouverture." && r.Range.Text == null && r.Get<Button>("adviceReformulate").Enabled && r.Get<Label>("status").Text.StartsWith("Avis reçu");
            } });
            await Test("Reformuler selon l'avis : avis transmis, reformulation dans l'onglet Reformulé", async () => { using (var r = new Rig()) {
                await r.Tick(); var advice = (Task)r.Call("RequestAdvice"); r.Runner.Pending.SetResult("Plus ferme et plus court."); await advice;
                r.Runner.Next(); r.Get<Button>("adviceReformulate").PerformClick(); r.Runner.Pending.SetResult("Bonjour, voici les documents. Merci de confirmer."); await Task.Delay(100);
                return r.Runner.Fluent && r.Runner.LastAdvice == "Plus ferme et plus court." && r.Get<TabControl>("textTabs").SelectedTab == r.Get<TabPage>("reformulationPage") && r.Get<TextBox>("reformulated").Text.StartsWith("Bonjour, voici") && r.Range.Text == null;
            } });
            await Test("Avis périmé après modification du mail : Reformuler selon l'avis grisé", async () => { using (var r = new Rig()) {
                await r.Tick(); var advice = (Task)r.Call("RequestAdvice"); r.Runner.Pending.SetResult("Plus court."); await advice;
                r.Text += "Merci"; r.Observe();
                return !r.Get<Button>("adviceReformulate").Enabled;
            } });
            await Test("Moteur réel de processus : Avis et reformulation selon l'avis transmis à l'IA", async () => { using (var engine = ProcessRunner()) {
                string advice = await engine.Advise("FAKE_PAYLOAD", "Trop gentil ?"); string follow = await engine.Correct("FAKE_PAYLOAD", true, "Plus ferme"); string strict = await engine.Correct("FAKE_PAYLOAD", false);
                return advice.Contains("\"mode\":\"AVIS\"") && advice.Contains("\"question\":\"Trop gentil ?\"") && follow.Contains("\"mode\":\"REFORMULER\"") && follow.Contains("\"avis\":\"Plus ferme\"") && strict.Contains("\"mode\":\"STRICT\"") && !strict.Contains("\"avis\"");
            } });
            await Test("Reformuler utilise le mode explicite", async () => { using (var r = new Rig()) { var t = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour."); await t; return r.Runner.Fluent && r.Valid; } });
            await Test("Onglet Reformulé : aucune demande, seul le bouton Reformuler reformule", async () => { using (var r = new Rig()) {
                await r.Ready(); r.Runner.Next();
                var tabs = r.Get<TabControl>("textTabs"); tabs.SelectedTab = r.Get<TabPage>("reformulationPage");
                await Task.Delay(50);
                bool idle = r.Runner.Calls == 1 && !r.Get<bool>("busy") && r.Get<Label>("status").Text.Contains("Reformuler") && r.Get<Button>("reformulate").Enabled;
                var task = (Task)r.Call("Request", true); r.Runner.Pending.SetResult("Bonjour, les documents sont désormais prêts."); await task;
                tabs.SelectedTab = r.Get<TabPage>("resultPage"); tabs.SelectedTab = r.Get<TabPage>("reformulationPage");
                return idle && r.Valid && r.Runner.Calls == 2 && r.Runner.Fluent && r.Get<TextBox>("reformulated").Text.Contains("désormais");
            } });
            await Test("Commandes compactes : Copier en bas, Pause près des réglages, Aperçu retiré", () => { using (var r = new Rig()) {
                var copy = r.Get<Button>("copy"); var pause = r.Get<Button>("pause"); var ai = r.Get<Button>("ai");
                return Value(copy.Parent == r.Get<Button>("retry").Parent && pause.Parent == ai.Parent && pause.Text == "Ⅱ" && r.Get<Button>("expand").Parent == null);
            } });
            await Test("Quitter réellement pendant une requête", async () => { using (var r = new Rig()) { var t = r.Request(); r.Window.ExitApplication(); r.Runner.Pending.SetResult("Réponse tardive"); await t; return r.Window.IsDisposed && r.Runner.Stops > 0; } });
            await Test("Brouillon absent efface la proposition", async () => { using (var r = new Rig()) { await r.Tick(); await r.Ready(); r.Exists = false; await r.Tick(); return !r.Valid && r.Get<TextBox>("result").Text.Length == 0; } });
            await Test("Changer entre 100 brouillons ne mélange pas les réponses", async () => { using (var r = new Rig()) { var t = r.Request(); for (int i = 0; i < 100; i++) { r.Key = "mail-" + i; r.Text = "Brouillon " + i; r.Observe(); } r.Runner.Pending.SetResult("Ancienne réponse"); await t; return !r.Valid && r.Get<TextBox>("manualInput").Text == "Brouillon 99"; } });
            await Test("100 corrections manuelles successives", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); for (int i = 0; i < 100; i++) { r.Get<TextBox>("manualInput").Text = "Texte " + i; if (i > 0) r.Runner.Next(); await r.Ready("Texte corrigé " + i); if (!r.Valid) return false; } return r.Runner.Calls == 100; } });
            await Test("Fenêtre redimensionnable sans boutons masqués", () => { using (var r = new Rig()) { r.ClientSize = new Size(640, 500); r.Window.PerformLayout(); var button = r.Get<Button>("retry"); return Value(button.Width >= 95 * r.DpiScale && button.Height >= 28 * r.DpiScale && r.Window.MinimumSize.Width >= 440); } });
            await Test("Accueil et boutons gardent leurs proportions à plusieurs tailles", () => { using (var r = new Rig()) { foreach (Size size in new [] { new Size(480, 160), new Size(520, 160), new Size(640, 500) }) { r.ClientSize = size; r.Window.PerformLayout(); if (r.Get<TabControl>("textTabs").Visible) return Value(false); foreach (string name in new [] { "ai", "theme", "copy", "retry", "apply", "pause", "manual" }) { var b = r.Get<Button>(name); var rect = r.Window.RectangleToClient(b.RectangleToScreen(b.ClientRectangle)); if (b.Height < 28 * r.DpiScale || b.Height > 55 * r.DpiScale || b.Width < 28 * r.DpiScale || !r.Window.ClientRectangle.Contains(rect)) return Value(false); } } return Value(true); } });
            await Test("L'aperçu et une réponse ne réduisent pas la fenêtre choisie", async () => { using (var r = new Rig()) { r.ClientSize = new Size(800, 650); await r.Tick(); bool preview = r.ClientSize == new Size(800, 650) && r.Get<TextBox>("manualInput").Height > 250 * r.DpiScale && r.Get<TabControl>("textTabs").Visible; await r.Ready(); return preview && r.ClientSize == new Size(800, 650) && r.Get<Button>("copy").Height <= 55 * r.DpiScale; } });
            await Test("Retour Outlook puis Manuel conserve les dimensions et la saisie", () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Mon texte conservé"; r.ClientSize = new Size(800, 650); r.Call("SwitchMode"); bool compact = !r.Get<TabControl>("textTabs").Visible && r.ClientSize.Height <= 240; r.Call("SwitchMode"); return Value(compact && r.ClientSize == new Size(800, 650) && r.Get<TextBox>("manualInput").Text == "Mon texte conservé" && r.Get<TabControl>("textTabs").Visible); } });
            await Test("Mise à l'échelle simulée à 150 et 200 % : actions contenues", () => { foreach (float scale in new [] { 1.5f, 2f }) { using (var r = new Rig()) { r.Call("SwitchMode"); r.Window.Scale(new SizeF(scale, scale)); r.Window.PerformLayout(); foreach (string name in new [] { "ai", "theme", "copy", "retry", "apply", "pause" }) { var b = r.Get<Button>(name); var rect = r.Window.RectangleToClient(b.RectangleToScreen(b.ClientRectangle)); if (!r.Window.ClientRectangle.Contains(rect) || b.Height > 55 * scale * r.DpiScale || b.Height < 28 * scale * r.DpiScale) return Value(false); } if (r.Get<TextBox>("manualInput").Height < 110 * scale * r.DpiScale) return Value(false); } } return Value(true); });
            await Test("Moteur réel de processus : JSON incorrect suivi d'une réponse UTF8", async () => { using (var engine = ProcessRunner()) return await engine.Correct("FAKE_NORMAL", false) == "Réponse fictive, prête 😊."; });
            await Test("Réponse affichable avant la fin du processus", async () => { using (var engine = ProcessRunner()) { var watch = Stopwatch.StartNew(); var answer = await engine.Correct("FAKE_EARLY", false); return answer.Contains("prête") && watch.ElapsedMilliseconds < 3000; } });
            await Test("Processus sans réponse produit une erreur exploitable", async () => { using (var engine = ProcessRunner()) return await Error(engine.Correct("FAKE_EMPTY", false), "exploitable"); });
            await Test("Authentification expirée signalée", async () => { using (var engine = ProcessRunner()) return await Error(engine.Correct("FAKE_AUTH", false), "renouveler"); });
            await Test("Quota épuisé signalé", async () => { using (var engine = ProcessRunner()) return await Error(engine.Correct("FAKE_QUOTA", false), "Limite"); });
            await Test("Délai dépassé arrête le processus", async () => { using (var engine = ProcessRunner(350)) { var watch = Stopwatch.StartNew(); bool valid = await Error(engine.Correct("FAKE_HANG", false), "Aucune correction"); return valid && watch.ElapsedMilliseconds < 3000; } });
            await Test("Annulation avant démarrage du processus", async () => { using (var engine = ProcessRunner()) { var task = engine.Correct("FAKE_HANG", false); engine.Stop(); return await Error(task, "annulée"); } });
            await Test("Annulation pendant exécution du processus", async () => { using (var engine = ProcessRunner()) { var task = engine.Correct("FAKE_HANG", false); await Task.Delay(250); engine.Stop(); return await Error(task, "annulée"); } });
            await Test("Réessai moteur après annulation", async () => { using (var engine = ProcessRunner()) { var task = engine.Correct("FAKE_HANG", false); await Task.Delay(250); engine.Stop(); if (!await Error(task, "annulée")) return false; return (await engine.Correct("FAKE_NORMAL", false)).Contains("prête"); } });
            await Test("Codex absent : message de connexion", async () => { using (var engine = ProcessRunner(3000, true)) return await Error(engine.Correct("FAKE_NORMAL", false), "introuvable"); });
            await TestAi();
            await TestThemes();
            await TestSettings();
        }
        private static void DeleteFixtureDirectory(string folder) {
            string safe = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TEMP")) + Path.DirectorySeparatorChar;
            if (!Path.GetFullPath(folder).StartsWith(safe, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Dossier de test hors TEMP.");
            var watch=Stopwatch.StartNew();
            while (Directory.Exists(folder)) {
                try { Directory.Delete(folder,true); return; }
                catch (IOException) { if(watch.ElapsedMilliseconds>2000) throw; Thread.Sleep(30); }
            }
        }
        private static bool HasBrandColor(Icon icon) { using (var bitmap = icon.ToBitmap()) return HasBrandColor(bitmap); }
        private static bool HasBrandColor(Bitmap bitmap) {
            for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++) {
                Color pixel = bitmap.GetPixel(x, y); if (pixel.A > 128 && pixel.G > pixel.R + 60 && pixel.B > pixel.R + 60 && Math.Abs(pixel.G - pixel.B) < 30) return true;
            } return false;
        }
        private static async Task TestSettings() {
            using (var small = AppBrand.LoadIcon(16)) using (var large = AppBrand.LoadIcon(32)) {
                Check("Icônes Korecto 16 et 32 pixels bleu-vert", small.Width == 16 && large.Width == 32 && HasBrandColor(small) && HasBrandColor(large));
            }
            using (var associated = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location)) Check("Icône Korecto intégrée au fichier exécutable", associated != null && HasBrandColor(associated));
            using (var notification = AppBrand.LoadNotificationIcon(16)) Check("Icône de notification Korecto dédiée 16 pixels", notification.Width == 16 && HasBrandColor(notification));
            using (var logo = AppBrand.LoadLogo()) Check("Logo Korecto de l'accueil intégré", logo.Width == 128 && HasBrandColor((Bitmap)logo));
            using (var r = new Rig()) Check("Fenêtre et notification utilisent l'identité Korecto", r.Window.Icon.Width == 32 && HasBrandColor(r.Window.Icon) && r.Get<NotifyIcon>("tray").Icon.Width == 16 && HasBrandColor(r.Get<NotifyIcon>("tray").Icon));
            var store = new FakeStartupStore(); var startup = new StartupPreference(store, @"C:\Mes applications\CorrecteurMail.exe");
            Check("Démarrage automatique désactivé par défaut", !startup.Enabled && store.Writes == 0);
            startup.Save(false); Check("Enregistrer sans activer ne modifie pas Windows", store.Writes == 0 && store.Removes == 0);
            startup.Save(true); Check("Activation : chemin avec espaces correctement protégé", store.Value == "\"C:\\Mes applications\\CorrecteurMail.exe\" --startup" && startup.Enabled);
            startup.Save(true); Check("Activation répétée ne crée pas de doublon", store.Writes == 1);
            startup.Save(false); Check("Désactivation enlève seulement l’entrée de l’application", store.Value == null && store.Removes == 1 && !startup.Enabled);
            store.Value = "\"C:\\Ancienne version\\CorrecteurMail.exe\" --startup";
            startup.Save(true); Check("Nouvelle version : enregistrement du chemin courant", store.Value == startup.Command && store.Writes == 2);
            await Test("Paramètres : cocher puis annuler ne change pas le démarrage", () => {
                var isolated = new FakeStartupStore(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => null, new StartupPreference(isolated))) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.Startup.Checked = true; dialog.DialogResult = DialogResult.Cancel; dialog.Close();
                    return Value(isolated.Writes == 0 && isolated.Removes == 0 && isolated.Value == null);
                }
            });
            await Test("Paramètres : cocher puis enregistrer active le démarrage", () => {
                var isolated = new FakeStartupStore(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => null, new StartupPreference(isolated))) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.Startup.Checked = true; dialog.Save.PerformClick();
                    return Value(isolated.Writes == 1 && isolated.Value.EndsWith(" --startup") && dialog.DialogResult == DialogResult.OK);
                }
            });
            await Test("Paramètres : accès refusé signalé sans activer le démarrage", () => {
                var isolated = new FakeStartupStore { DenyWrite = true }; using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => null, new StartupPreference(isolated))) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.Startup.Checked = true; dialog.Save.PerformClick();
                    return Value(isolated.Value == null && dialog.DialogResult != DialogResult.OK && dialog.Visible);
                }
            });
            await Test("Paramètres : choix de l’IA disponible sans modifier le démarrage", () => {
                int calls = 0; var isolated = new FakeStartupStore(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => { calls++; return null; }, new StartupPreference(isolated))) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.AiChoice.PerformClick();
                    return Value(calls == 1 && isolated.Writes == 0 && dialog.Startup.Checked == false);
                }
            });
            await Test("Lancement Windows discret, puis affichage à la demande", () => {
                using (var runner = new FakeRunner()) using (var window = new Popup(true, runner, () => null, () => false, startInTray: true)) {
                    window.ShowInTaskbar = false; window.Opacity = 0; window.Show(); bool hidden = !window.Visible && window.IsHandleCreated;
                    window.Show(); return Value(hidden && window.Visible && runner.Calls == 0);
                }
            });
            await Test("Croix : masque la fenêtre et conserve le texte manuel", () => {
                using (var r = new Rig()) {
                    r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Texte à conserver"; r.Window.Close();
                    return Value(!r.Window.IsDisposed && !r.Window.Visible && !r.Window.ShowInTaskbar && r.Get<TextBox>("manualInput").Text == "Texte à conserver" && r.Runner.Stops == 0);
                }
            });
            await Test("Fenêtre masquée : réouverture automatique à la prochaine saisie Outlook", async () => {
                using (var r = new Rig()) { await r.Tick(); r.Window.Close(); int reads = r.Reads; r.Text = "Brouillon modifié pendant la réduction"; await r.Tick(); return r.Window.Visible && r.Reads > reads && r.Runner.Calls == 0 && r.Get<TextBox>("manualInput").Text == r.Text; }
            });
            await Test("Croix sans nouvelle saisie ne fait pas réapparaître la fenêtre", async () => {
                using (var r = new Rig()) { await r.Tick(); r.Window.Close(); await r.Tick(); return !r.Window.Visible && r.Runner.Calls == 0; }
            });
            await Test("Retour sur Outlook fait apparaître Korecto sans clic", async () => {
                using (var r = new Rig()) { r.Window.Hide(); r.Foreground = false; await r.Tick(); bool hidden = !r.Window.Visible; r.Foreground = true; await r.Tick(); if (!(hidden && r.Window.Visible && r.Get<TextBox>("manualInput").Text.Contains("Bonjour"))) throw new Exception("hidden=" + hidden + "; visible=" + r.Window.Visible + "; text=" + r.Get<TextBox>("manualInput").Text + "; focus=" + r.Window.ContainsFocus); return true; }
            });
            await Test("Afficher depuis la notification restaure l’aperçu", async () => {
                using (var r = new Rig()) {
                    await r.Tick(); r.Window.Close(); r.Text = "Nouveau texte après restauration";
                    foreach (ToolStripItem item in r.Get<NotifyIcon>("tray").ContextMenuStrip.Items) if (item.Text == "Afficher") item.PerformClick();
                    await r.Tick(); return r.Window.Visible && r.Window.ShowInTaskbar && r.Get<TextBox>("manualInput").Text == r.Text && r.Runner.Calls == 0;
                }
            });
            await Test("Croix pendant correction : résultat conservé sans réapparition", async () => {
                using (var r = new Rig()) {
                    r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Les document sont prêt."; var task = r.Request(); r.Window.Close(); r.Runner.Pending.SetResult("Les documents sont prêts."); await task;
                    bool retained = r.Valid && !r.Window.Visible && r.Runner.Stops == 0;
                    r.Window.RestoreFromTray(); return retained && r.Window.Visible && r.Get<TextBox>("result").Text == "Les documents sont prêts.";
                }
            });
            await Test("Quitter dans le menu arrête réellement l’application", () => {
                using (var r = new Rig()) {
                    r.Window.Close();
                    foreach (ToolStripItem item in r.Get<NotifyIcon>("tray").ContextMenuStrip.Items) if (item.Text == "Quitter") { item.PerformClick(); break; }
                    return Value(r.Window.IsDisposed && r.Runner.Stops > 0);
                }
            });
            await Test("La fermeture de session Windows n’est pas bloquée", () => {
                using (var r = new Rig()) { var args = new FormClosingEventArgs(CloseReason.WindowsShutDown, false); r.Call("OnFormClosing", args); return Value(!args.Cancel); }
            });
        }
        private static async Task TestThemes() {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TEMP", "THEME", "preferences.theme.txt");
            AppearanceStore.Save(false, path); Check("Thème clair conservé pour le prochain lancement", !AppearanceStore.LoadDark(path));
            AppearanceStore.Save(true, path); Check("Retour au thème sombre conservé", AppearanceStore.LoadDark(path));
            Check("Thème sombre par défaut sans préférence", AppearanceStore.LoadDark(path + ".absent"));
            File.WriteAllText(path, "valeur invalide"); Check("Préférence illisible revient au thème sombre", AppearanceStore.LoadDark(path));
            bool original = Palette.IsDark;
            try {
                Palette.Use(true);
                await Test("Passer au clair conserve la correction et ses actions", async () => { using (var r = new Rig()) { await r.Ready(); string correction = r.Get<TextBox>("result").Text; long version = r.Get<Debounce>("state").Version; r.Get<Button>("theme").PerformClick(); return !Palette.IsDark && r.Window.BackColor.GetBrightness() > 0.8 && r.Get<TextBox>("result").BackColor == Color.White && r.Get<TextBox>("result").Text == correction && r.Valid && r.Get<Debounce>("state").Version == version && r.Runner.Calls == 1; } });
                Palette.Use(true);
                await Test("Changer de thème pendant la correction préserve la saisie et la requête", async () => { using (var r = new Rig()) { r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Bonjour, les document sont prêt."; var task = r.Request(); r.Get<Button>("theme").PerformClick(); bool waiting = r.Get<bool>("busy") && r.Runner.Stops == 0; r.Runner.Pending.SetResult("Bonjour, les documents sont prêts."); await task; return waiting && r.Valid && r.Get<TextBox>("manualInput").Text == "Bonjour, les document sont prêt." && r.Runner.Calls == 1; } });
                Palette.Use(true);
                await Test("Le mode manuel ouvert après le passage au clair reçoit le bon thème", () => { using (var r = new Rig()) { r.Get<Button>("theme").PerformClick(); r.Call("SwitchMode"); return Value(r.Get<TextBox>("manualInput").BackColor == Color.White && r.Get<TextBox>("manualInput").ForeColor == Palette.Text && r.Get<TabPage>("inputPage").BackColor == Color.White); } });
                Palette.Use(true);
                await Test("Choix d'IA : thème clair puis sombre sans changer le modèle", () => { using (var dialog = new AiSettingsDialog(AiSettings.Default(), autoCheck: false, appearancePath: path)) { dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.SetConnections(new [] { new AiConnection { Provider = "chatgpt", Connected = true } }); dialog.ThemeChoice.PerformClick(); bool light = dialog.BackColor.GetBrightness() > 0.8 && dialog.Models.BackColor == Color.White && dialog.Chatgpt.Checked && dialog.Models.SelectedIndex == 0; dialog.ThemeChoice.PerformClick(); return Value(light && Palette.IsDark && dialog.BackColor.GetBrightness() < 0.2 && AppearanceStore.LoadDark(path) && dialog.Models.SelectedIndex == 0); } });
            } finally { Palette.Use(original); }
        }
        private static async Task TestAi() {
            string directory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TEMP", "IA"); Directory.CreateDirectory(directory);
            string config = Path.Combine(directory, "parametres.json");
            Check("Luna recommandé par défaut", AiSettings.Default().Model == "gpt-6-luna");
            Check("Haiku recommandé par défaut", AiSettings.Default("claude").Model == "haiku");
            Check("Rapport du test réel : modèle réellement utilisé (Luna, pas Terra)", Program.TestModelLine(AiSettings.Default()).Contains("gpt-6-luna") && !Program.TestModelLine(AiSettings.Default()).Contains("terra"));
            Check("Rapport du test réel : Claude/Haiku sans mention d'effort ChatGPT", Program.TestModelLine(AiSettings.Default("claude")).Contains("Claude / haiku") && !Program.TestModelLine(AiSettings.Default("claude")).Contains("effort"));
            Check("Modèle d'un autre fournisseur refusé", !new AiSettings { Provider = "claude", Model = "gpt-6.1-sol" }.Valid);
            Check("Fournisseur inconnu refusé", !new AiSettings { Provider = "autre", Model = "haiku" }.Valid);
            Check("Premier lancement sans fichier impose un choix", AiSettingsStore.Load(Path.Combine(directory, "inexistant.json")) == null);
            AiSettingsStore.Save(AiSettings.Default(), config); var loaded = AiSettingsStore.Load(config);
            Check("Choix ChatGPT/Luna conservé au redémarrage", loaded != null && loaded.Provider == "chatgpt" && loaded.Model == "gpt-6-luna");
            AiSettingsStore.Save(new AiSettings { Provider = "claude", Model = "sonnet" }, config); loaded = AiSettingsStore.Load(config);
            Check("Remplacement des paramètres conservé", loaded != null && loaded.Provider == "claude" && loaded.Model == "sonnet");
            Check("Paramètres sans clé ni identité de compte", File.ReadAllText(config) == "{\"Provider\":\"claude\",\"Model\":\"sonnet\"}");
            File.WriteAllText(config, "{invalid"); Check("Paramètres corrompus redemandent le choix", AiSettingsStore.Load(config) == null);
            File.WriteAllText(config, "{\"Provider\":\"chatgpt\",\"Model\":\"inconnu\"}"); Check("Ancien modèle inconnu redemande le choix", AiSettingsStore.Load(config) == null);
            Check("Connexion ChatGPT locale reconnue", LocalAi.IsAccountLogin("chatgpt", 0, "Logged in using ChatGPT"));
            Check("Connexion Codex par API refusée", !LocalAi.IsAccountLogin("chatgpt", 0, "Logged in using an API key"));
            Check("Codex non connecté refusé", !LocalAi.IsAccountLogin("chatgpt", 1, "Logged in using ChatGPT"));
            Check("Compte Claude local reconnu", LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":true,\"authMethod\":\"claude.ai\"}"));
            Check("Clé API Claude refusée", !LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":true,\"authMethod\":\"api_key\"}"));
            Check("Fournisseur cloud Claude refusé", !LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":true,\"authMethod\":\"third_party\"}"));
            Check("Token fourni par environnement Claude refusé", !LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":true,\"authMethod\":\"oauth_token\"}"));
            Check("Claude déconnecté refusé", !LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":false,\"authMethod\":\"claude.ai\"}"));
            Check("État Claude mal formé refusé", !LocalAi.IsAccountLogin("claude", 0, "{\"loggedIn\":42}"));
            var command = new LocalCommand { FileName = Assembly.GetExecutingAssembly().Location };
            var start = command.StartInfo("exec", directory);
            Check("Aucune clé API ni token alternatif transmis", !start.EnvironmentVariables.ContainsKey("OPENAI_API_KEY") && !start.EnvironmentVariables.ContainsKey("ANTHROPIC_API_KEY") && !start.EnvironmentVariables.ContainsKey("CLAUDE_CODE_OAUTH_TOKEN") && !start.EnvironmentVariables.ContainsKey("ANTHROPIC_BASE_URL"));
            Check("Sonde réelle de processus Codex simulé", LocalAi.Check("chatgpt", command, directory).Connected);
            Check("Sonde réelle de processus Claude simulé", LocalAi.Check("claude", command, directory).Connected);
            var both = new [] { new AiConnection { Provider = "chatgpt", Connected = true, Message = "Compte connecté" }, new AiConnection { Provider = "claude", Connected = true, Message = "Compte connecté" } };
            using (var dialog = new AiSettingsDialog(null, config, autoCheck: false)) {
                dialog.SetConnections(both); dialog.Claude.Checked = true; dialog.Models.SelectedIndex = 0;
                Check("Premier écran permet de choisir Claude/Haiku", dialog.SaveSelection() && AiSettingsStore.Load(config).Model == "haiku");
            }
            using (var dialog = new AiSettingsDialog(new AiSettings { Provider = "chatgpt", Model = "gpt-6.1-sol" }, config, autoCheck: false)) {
                dialog.SetConnections(both);
                Check("Paramètres retrouvent le modèle choisi", dialog.Chatgpt.Checked && dialog.Models.SelectedIndex == 2);
                dialog.Claude.Checked = true; dialog.Models.SelectedIndex = 1;
                Check("Changement d'IA et de modèle enregistrable", dialog.SaveSelection() && dialog.Selected.Provider == "claude" && dialog.Selected.Model == "sonnet");
                dialog.SetConnections(new [] { new AiConnection { Provider = "chatgpt", Connected = false }, new AiConnection { Provider = "claude", Connected = false } });
                Check("Aucune IA connectée : enregistrement bloqué", !dialog.Accept.Enabled && !dialog.SaveSelection() && !dialog.Chatgpt.Enabled && !dialog.Claude.Enabled);
            }
            using (var dialog = new AiSettingsDialog(null, config, autoCheck: false)) {
                dialog.SetConnections(new [] { both[0], new AiConnection { Provider = "claude", Connected = false } });
                Check("Claude non connecté ne peut pas être choisi", !dialog.Claude.Enabled && dialog.Chatgpt.Enabled);
                dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.PerformLayout();
                Check("Modèle et validation lisibles dans une fenêtre compacte", dialog.Models.SelectedIndex == 0 && dialog.Models.DrawMode == DrawMode.OwnerDrawFixed && dialog.Accept.Height >= 30 * DisplayScale(dialog) && dialog.Accept.Height <= 55 * DisplayScale(dialog) && dialog.Accept.Bottom <= dialog.Accept.Parent.ClientSize.Height);
            }
            using (var dialog = new AiSettingsDialog(new AiSettings { Provider = "chatgpt", Model = "gpt-6.1-sol" }, config, autoCheck: false)) {
                dialog.SetConnections(both); dialog.Claude.Checked = true;
                bool claudeCheapest = dialog.Models.SelectedIndex == 0 && Convert.ToString(dialog.Models.Items[0]).StartsWith("Haiku");
                dialog.Chatgpt.Checked = true;
                Check("Changement de fournisseur : Haiku présélectionné, modèle enregistré retrouvé au retour", claudeCheapest && dialog.Models.SelectedIndex == 2);
            }
            using (var dialog = new AiSettingsDialog(null, config, autoCheck: false)) {
                dialog.SetConnections(both); bool luna = Convert.ToString(dialog.Models.SelectedItem).StartsWith("Luna");
                dialog.Claude.Checked = true;
                Check("Premier choix : Luna puis Haiku proposés d'office", luna && Convert.ToString(dialog.Models.SelectedItem).StartsWith("Haiku"));
            }
            string pendingConfig = Path.Combine(directory, "non-enregistre.json"); if (File.Exists(pendingConfig)) File.Delete(pendingConfig);
            using (var dialog = new AiSettingsDialog(AiSettings.Default(), pendingConfig, autoCheck: false, persist: false)) {
                dialog.SetConnections(both); dialog.Claude.Checked = true; dialog.Models.SelectedIndex = 0;
                Check("Choix d'IA depuis Paramètres : rien enregistré avant Enregistrer", dialog.SaveSelection() && dialog.Selected.Model == "haiku" && !File.Exists(pendingConfig) && dialog.Accept.Text == "Valider");
            }
            await Test("Paramètres : IA choisie puis Annuler garde l'IA actuelle", () => {
                var applied = new List<AiSettings>(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => AiSettings.Default("claude"), new StartupPreference(new FakeStartupStore()), applyAi: applied.Add)) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.AiChoice.PerformClick(); dialog.DialogResult = DialogResult.Cancel; dialog.Close();
                    return Value(applied.Count == 0 && dialog.PendingAi != null);
                }
            });
            await Test("Paramètres : IA choisie puis Enregistrer l'applique une fois", () => {
                var applied = new List<AiSettings>(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => AiSettings.Default("claude"), new StartupPreference(new FakeStartupStore()), applyAi: applied.Add)) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.AiChoice.PerformClick(); dialog.Save.PerformClick();
                    return Value(applied.Count == 1 && applied[0].Provider == "claude" && applied[0].Model == "haiku" && dialog.DialogResult == DialogResult.OK);
                }
            });
            await Test("Paramètres : même IA rechoisie, rien à appliquer", () => {
                var applied = new List<AiSettings>(); using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => AiSettings.Default(), new StartupPreference(new FakeStartupStore()), applyAi: applied.Add)) {
                    dialog.Opacity = 0; dialog.ShowInTaskbar = false; dialog.Show(); dialog.AiChoice.PerformClick(); dialog.Save.PerformClick();
                    return Value(applied.Count == 0 && dialog.PendingAi == null && dialog.DialogResult == DialogResult.OK);
                }
            });
            using (var r = new Rig()) {
                var task = r.Request(); var replacement = new FakeRunner(); r.Window.ChangeAi(AiSettings.Default("claude"), replacement); r.Runner.Pending.SetResult("Ancienne réponse ChatGPT"); await task;
                Check("Changer d'IA rejette la réponse de l'ancienne IA", !r.Valid && r.Window.Text.Contains("Claude") && r.Runner.Stops > 0);
            }
            using (var r = new Rig()) {
                r.Call("SwitchMode"); r.Get<TextBox>("manualInput").Text = "Mon texte à garder"; await r.Ready("Mon texte corrigé"); r.Window.ChangeAi(AiSettings.Default("claude"), new FakeRunner());
                Check("Changer d'IA conserve la saisie et invalide le résultat", r.Get<TextBox>("manualInput").Text == "Mon texte à garder" && !r.Valid);
            }
            var parser = new JavaScriptSerializer();
            Check("Claude : réponse structurée finale acceptée", AccountRunner.ReadClaudeCorrection("{\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"structured_output\":{\"texte\":\"Été 😊\"}}", parser) == "Été 😊");
            Check("Claude : message assistant intermédiaire ignoré", AccountRunner.ReadClaudeCorrection("{\"type\":\"assistant\",\"texte\":\"Brouillon\"}", parser) == null);
            Check("Claude : résultat en erreur ignoré", AccountRunner.ReadClaudeCorrection("{\"type\":\"result\",\"is_error\":true,\"structured_output\":{\"texte\":\"Erreur\"}}", parser) == null);
            Check("Claude : réponse vide refusée", AccountRunner.ReadClaudeCorrection("{\"type\":\"result\",\"is_error\":false,\"structured_output\":{\"texte\":\" \"}}", parser) == null);
            Check("Claude : résultat mal typé ignoré", AccountRunner.ReadClaudeCorrection("{\"type\":42}", parser) == null);
            await Test("Claude : lancement et réception Haiku simulés", async () => { using (var engine = ProcessRunner(settings: AiSettings.Default("claude"))) return (await engine.Correct("FAKE_NORMAL", false)).Contains("prête"); });
            await Test("Claude : schéma JSON et outils vides correctement échappés", async () => { using (var engine = ProcessRunner(settings: AiSettings.Default("claude"))) return await engine.Correct("FAKE_SCHEMA", false) == "Schéma et outils vérifiés"; });
            await Test("Luna explicitement transmis au client Codex", async () => { using (var engine = ProcessRunner()) { string args = await engine.Correct("FAKE_ARGUMENTS", false); return args.Contains("--model|gpt-6-luna") && args.Contains("model_reasoning_effort=\"low\"") && !args.Contains("FAKE_ARGUMENTS"); } });
            await Test("Sol explicitement transmis au client Codex", async () => { using (var engine = ProcessRunner(settings: new AiSettings { Provider = "chatgpt", Model = "gpt-6.1-sol" })) return (await engine.Correct("FAKE_ARGUMENTS", false)).Contains("--model|gpt-6.1-sol"); });
            await Test("Sonnet explicitement transmis au client Claude", async () => { using (var engine = ProcessRunner(settings: new AiSettings { Provider = "claude", Model = "sonnet" })) return (await engine.Correct("FAKE_ARGUMENTS", false)).Contains("--model|sonnet"); });
            await Test("Compte déconnecté bloque la correction sans repli", async () => { using (var engine = ProcessRunner(connected: false)) return await Error(engine.Correct("FAKE_NORMAL", false), "déconnecté"); });
            await Test("Claude : erreur de quota correctement signalée", async () => { using (var engine = ProcessRunner(settings: AiSettings.Default("claude"))) return await Error(engine.Correct("FAKE_QUOTA", false), "Claude"); });
            await Test("Client fermé avant lecture de la demande : message de mise à jour", async () => {
                Environment.SetEnvironmentVariable("KORECTO_FAKE_EARLY_EXIT", "1");
                try { using (var engine = ProcessRunner()) return await Error(engine.Correct(new string('a', 200000), false), "incompatible"); }
                finally { Environment.SetEnvironmentVariable("KORECTO_FAKE_EARLY_EXIT", null); }
            });
            Check("Réponse intégrée : identité stable malgré un nouvel objet COM", Draft.InlineIdentity("0101AB", 111) == Draft.InlineIdentity("0101AB", 222));
            string fakeClient = Assembly.GetExecutingAssembly().Location, fakeRoot = Path.GetDirectoryName(Path.GetDirectoryName(fakeClient));
            await Test("Connexion vérifiée une seule fois pour plusieurs corrections", async () => {
                int checks = 0; using (var engine = new AccountRunner(() => fakeClient, fakeRoot, 3000, null, () => { checks++; return new AiConnection { Connected = true }; })) {
                    bool ok = (await engine.Correct("FAKE_NORMAL", false)).Contains("prête") && (await engine.Correct("FAKE_NORMAL", false)).Contains("prête");
                    return ok && checks == 1;
                }
            });
            await Test("Échec de correction : connexion revérifiée à la suivante", async () => {
                int checks = 0; using (var engine = new AccountRunner(() => fakeClient, fakeRoot, 3000, null, () => { checks++; return new AiConnection { Connected = true }; })) {
                    await engine.Correct("FAKE_NORMAL", false);
                    if (!await Error(engine.Correct("FAKE_AUTH", false), "renouveler")) return false;
                    await engine.Correct("FAKE_NORMAL", false); return checks == 2;
                }
            });
            await Test("Compte déconnecté : vérifié à chaque essai, jamais mis en confiance", async () => {
                int checks = 0; using (var engine = new AccountRunner(() => fakeClient, fakeRoot, 3000, null, () => { checks++; return new AiConnection { Connected = false, Message = "Compte déconnecté simulé" }; })) {
                    bool refused = await Error(engine.Correct("FAKE_NORMAL", false), "déconnecté") && await Error(engine.Correct("FAKE_NORMAL", false), "déconnecté");
                    return refused && checks == 2;
                }
            });
            var processNames = new Dictionary<uint, string> { { 10, "OUTLOOK" }, { 20, "Korecto" } };
            var foreground = new OutlookProcessCache(id => { if (!processNames.ContainsKey(id)) throw new ArgumentException(); return processNames[id]; });
            bool sameOutlook = foreground.IsOutlook(10) && foreground.IsOutlook(10) && foreground.IsOutlook(10);
            Check("Lecture Outlook : nom du processus au premier plan lu une seule fois", sameOutlook && foreground.Lookups == 1);
            Check("Lecture Outlook : nouvelle lecture seulement au changement de fenêtre", !foreground.IsOutlook(20) && foreground.IsOutlook(10) && foreground.Lookups == 3);
            Check("Lecture Outlook : processus disparu traité comme hors Outlook", !foreground.IsOutlook(99));
            Check("Réponse intégrée : deux réponses distinctes restent distinctes", Draft.InlineIdentity("0101AB", 1) != Draft.InlineIdentity("0101CD", 1) && Draft.InlineIdentity(null, 7) != Draft.InlineIdentity(null, 8) && Draft.InlineIdentity("", 7) == ":inline:7");
            string portable = Path.Combine(directory, "Dossier portable avec espaces"); Directory.CreateDirectory(portable);
            Check("Version portable : aucune consigne externe nécessaire", !File.Exists(Path.Combine(portable, "CONSIGNES.txt")) && AppAssets.Read(portable, "CONSIGNES.txt").Contains("MODE STRICT"));
            Check("Version portable : schéma intégré valide", new JavaScriptSerializer().DeserializeObject(AppAssets.Read(portable, "response.schema.json")) != null);
            string schemaFile = AppAssets.SchemaPath(portable);
            Check("Version portable : schéma extrait dans le profil utilisateur, rien à côté du programme", File.Exists(schemaFile) && schemaFile.StartsWith(AppFolders.Data) && !Directory.Exists(Path.Combine(portable, "PARAMETRES_IA")) && !File.Exists(Path.Combine(portable, "response.schema.json")));
            Check("Réglages : emplacement par défaut dans le dossier de données", AiSettingsStore.DefaultPath == Path.Combine(AppFolders.Data, "ia.json") && AppearanceStore.DefaultPath == Path.Combine(AppFolders.Data, "theme.txt"));
            TestMigration(directory);
            TestInstallation(directory);
            TestUpdates(directory);
            await Test("Version portable : correction Codex sans fichiers annexes", async () => {
                using (var engine = new AccountRunner(() => Assembly.GetExecutingAssembly().Location, portable, 3000, AiSettings.Default(), () => new AiConnection { Connected = true })) return (await engine.Correct("FAKE_NORMAL", false)).Contains("prête");
            });
            await Test("Version portable : correction Claude sans fichiers annexes", async () => {
                using (var engine = new AccountRunner(() => Assembly.GetExecutingAssembly().Location, portable, 3000, AiSettings.Default("claude"), () => new AiConnection { Connected = true })) return await engine.Correct("FAKE_SCHEMA", false) == "Schéma et outils vérifiés";
            });
        }
        private static void TestMigration(string directory) {
            string legacy = Path.Combine(directory, "ANCIEN"), data = Path.Combine(directory, "NOUVEAU");
            if (Directory.Exists(legacy)) Directory.Delete(legacy, true); if (Directory.Exists(data)) Directory.Delete(data, true);
            Directory.CreateDirectory(legacy);
            File.WriteAllText(Path.Combine(legacy, "PC-TEST.json"), "{\"Provider\":\"claude\",\"Model\":\"haiku\"}");
            File.WriteAllText(Path.Combine(legacy, "PC-TEST.theme.txt"), "clair");
            File.WriteAllText(Path.Combine(legacy, "AUTRE-PC.json"), "{\"Provider\":\"chatgpt\",\"Model\":\"gpt-6-luna\"}");
            Check("Migration : réglages du PC courant repris, ceux d'un autre PC ignorés", AppFolders.MigrateLegacy(legacy, data, "PC-TEST") == 2 && File.ReadAllText(Path.Combine(data, "ia.json")).Contains("haiku") && File.ReadAllText(Path.Combine(data, "theme.txt")) == "clair" && Directory.GetFiles(data).Length == 2);
            Check("Migration : originaux conservés", File.Exists(Path.Combine(legacy, "PC-TEST.json")) && File.Exists(Path.Combine(legacy, "PC-TEST.theme.txt")));
            File.WriteAllText(Path.Combine(data, "ia.json"), "{\"Provider\":\"chatgpt\",\"Model\":\"gpt-6.1-sol\"}");
            Check("Migration : réglages déjà présents jamais écrasés, rien de repris deux fois", AppFolders.MigrateLegacy(legacy, data, "PC-TEST") == 0 && File.ReadAllText(Path.Combine(data, "ia.json")).Contains("sol"));
            Check("Migration : dossier ancien absent sans erreur ni création", AppFolders.MigrateLegacy(Path.Combine(directory, "ABSENT"), Path.Combine(directory, "NOUVEAU2"), "PC-TEST") == 0 && !Directory.Exists(Path.Combine(directory, "NOUVEAU2")));
            var loaded = AiSettingsStore.Load(Path.Combine(data, "ia.json"));
            Check("Migration : réglages relus par l'application", loaded != null && loaded.Provider == "chatgpt" && loaded.Model == "gpt-6.1-sol");
        }
        internal sealed class FakeInstallRegistry : IInstallRegistry {
            internal string Run; internal Dictionary<string, object> Uninstall;
            public string ReadRun() { return Run; }
            public void WriteRun(string command) { Run = command; }
            public void RemoveRun() { Run = null; }
            public void WriteUninstall(Dictionary<string, object> values) { Uninstall = values; }
            public void RemoveUninstall() { Uninstall = null; }
        }
        private static void TestInstallation(string directory) {
            string root = Path.Combine(directory, "INSTALL"); if (Directory.Exists(root)) Directory.Delete(root, true);
            string download = Path.Combine(root, "Telechargements"), programs = Path.Combine(root, "Programs", "Korecto"), data = Path.Combine(root, "Data"), menu = Path.Combine(root, "Menu");
            Directory.CreateDirectory(Path.Combine(download, "PARAMETRES_IA"));
            string exe = Path.Combine(download, "Korecto.exe"); File.WriteAllText(exe, "exe-v1");
            string machine = String.Join("_", Environment.MachineName.Split(Path.GetInvalidFileNameChars()));
            File.WriteAllText(Path.Combine(download, "PARAMETRES_IA", machine + ".json"), "{\"Provider\":\"claude\",\"Model\":\"haiku\"}");
            var registry = new FakeInstallRegistry(); string deleted = null;
            var installer = new SelfInstaller(programs, data, menu, registry, folder => { deleted = folder; DeleteFixtureDirectory(folder); });
            Check("Installation : rien d'installé au départ, la copie téléchargée n'est pas la copie installée", !installer.Installed && !installer.IsInstalledCopy(exe));
            installer.Install(exe);
            Check("Installation : exe copié dans le dossier du programme", installer.Installed && File.ReadAllText(installer.InstalledExe) == "exe-v1" && installer.IsInstalledCopy(installer.InstalledExe));
            Check("Installation : raccourci du menu Démarrer créé", File.Exists(Path.Combine(menu, "Korecto.lnk")));
            Check("Installation : entrée de désinstallation Windows sans script", registry.Uninstall != null && (string)registry.Uninstall["UninstallString"] == "\"" + installer.InstalledExe + "\" --desinstaller" && (string)registry.Uninstall["DisplayName"] == "Korecto");
            Check("Installation : démarrage Windows non activé si il ne l'était pas", registry.Run == null);
            Check("Installation : anciens réglages de ce PC repris dans le profil", File.Exists(Path.Combine(data, "ia.json")) && File.ReadAllText(Path.Combine(data, "ia.json")).Contains("haiku"));
            File.WriteAllText(Path.Combine(programs, "DESINSTALLER.ps1"), "ancien");
            File.WriteAllText(Path.Combine(data, "ia.json"), "{\"Provider\":\"chatgpt\",\"Model\":\"gpt-6-luna\"}");
            File.WriteAllText(exe, "exe-v2"); registry.Run = "\"C:\\ancien\\Korecto.exe\" --startup";
            Check("Mise à jour : hash différent détecté", SelfInstaller.Hash(exe) != SelfInstaller.Hash(installer.InstalledExe));
            installer.Install(exe);
            Check("Mise à jour : nouvel exe, réglages conservés, anciens scripts retirés", File.ReadAllText(installer.InstalledExe) == "exe-v2" && File.ReadAllText(Path.Combine(data, "ia.json")).Contains("luna") && !File.Exists(Path.Combine(programs, "DESINSTALLER.ps1")) && SelfInstaller.Hash(exe) == SelfInstaller.Hash(installer.InstalledExe));
            Check("Mise à jour : démarrage Windows actif redirigé vers le programme installé", registry.Run == "\"" + installer.InstalledExe + "\" --startup");
            installer.Uninstall(false);
            Check("Désinstallation : programme, raccourci, démarrage et entrée Windows retirés", deleted == installer.ProgramsFolder && !installer.Installed && !File.Exists(Path.Combine(menu, "Korecto.lnk")) && registry.Run == null && registry.Uninstall == null);
            Check("Désinstallation : réglages conservés par défaut", File.Exists(Path.Combine(data, "ia.json")));
            installer.Install(exe); registry.Run = "\"D:\\autre\\Korecto.exe\" --startup";
            installer.Uninstall(true);
            Check("Désinstallation : démarrage pointant ailleurs jamais supprimé", registry.Run == "\"D:\\autre\\Korecto.exe\" --startup");
            Check("Désinstallation : réglages supprimés sur demande", !Directory.Exists(data));
            installer.RememberDeclined();
            Check("Installation refusée : mémorisée, plus de question", installer.Declined);
        }
        private static string ReleaseJson(string tag, string url, long size, string digest, bool draft = false, bool prerelease = false) {
            return "{\"tag_name\":\"" + tag + "\",\"draft\":" + (draft ? "true" : "false") + ",\"prerelease\":" + (prerelease ? "true" : "false") + ",\"assets\":[{\"name\":\"Autre.zip\",\"size\":1,\"browser_download_url\":\"https://github.com/mfarchitecture-ctrl/korecto/releases/download/" + tag + "/Autre.zip\"},{\"name\":\"Korecto.exe\",\"size\":" + size + ",\"digest\":" + (digest == null ? "null" : "\"" + digest + "\"") + ",\"browser_download_url\":\"" + url + "\"}]}";
        }
        private static string Sha(string file) { using (var sha = System.Security.Cryptography.SHA256.Create()) using (var s = File.OpenRead(file)) return "sha256:" + BitConverter.ToString(sha.ComputeHash(s)).Replace("-", "").ToLowerInvariant(); }
        private static void TestUpdates(string directory) {
            string root = Path.Combine(directory, "MAJ"); if (Directory.Exists(root)) Directory.Delete(root, true); Directory.CreateDirectory(root);
            string good = "https://github.com/mfarchitecture-ctrl/korecto/releases/download/v9.0.0/Korecto.exe";
            Check("Version : l'exe porte son numéro 1.0.2", AppBrand.Version == "1.0.2" && System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion == "1.0.2.0");
            Check("Version : étiquette v1.0.2 plus récente que 1.0.1, 1.0.1 et 1.0.0 non", Updater.IsNewer("1.0.1", Updater.ParseVersion("v1.0.2")) && !Updater.IsNewer("1.0.1", Updater.ParseVersion("v1.0.1")) && !Updater.IsNewer("1.0.1", Updater.ParseVersion("1.0.0")) && Updater.IsNewer("1.0.1", Updater.ParseVersion("v1.1")) && Updater.IsNewer("1.0.9", Updater.ParseVersion("1.10.0")));
            Check("Version : étiquettes invalides ignorées", Updater.ParseVersion("beta") == null && Updater.ParseVersion("") == null && Updater.ParseVersion(null) == null && !Updater.IsNewer("1.0.1", null) && !Updater.IsNewer("???", new Version(9, 0, 0)));
            var release = Updater.ParseRelease(ReleaseJson("v9.0.0", good, 2048, "sha256:ABCDEF"));
            Check("Publication : exe, taille, empreinte et version lus", release != null && release.Version == new Version(9, 0, 0) && release.AssetUrl == good && release.Size == 2048 && release.Digest == "sha256:ABCDEF");
            Check("Publication : brouillon, préversion, JSON absurde et exe manquant ignorés", Updater.ParseRelease(ReleaseJson("v9.0.0", good, 1, null, draft: true)) == null && Updater.ParseRelease(ReleaseJson("v9.0.0", good, 1, null, prerelease: true)) == null && Updater.ParseRelease("pas du json") == null && Updater.ParseRelease("{}") == null && Updater.ParseRelease("[1,2]") == null && Updater.ParseRelease("{\"tag_name\":\"v9.0.0\",\"assets\":[]}") == null);
            Check("Publication : adresse de téléchargement hors du dépôt refusée", Updater.ParseRelease(ReleaseJson("v9.0.0", "https://exemple.com/Korecto.exe", 1, null)) == null && Updater.ParseRelease(ReleaseJson("v9.0.0", "https://github.com/autre/depot/releases/download/v9/Korecto.exe", 1, null)) == null);
            Check("Publication : sans empreinte fournie, la release reste utilisable", Updater.ParseRelease(ReleaseJson("v9.0.0", good, 2048, null)).Digest == null);
            string file = Path.Combine(root, "fichier.exe"); var bytes = new byte[2048]; bytes[0] = (byte)'M'; bytes[1] = (byte)'Z'; File.WriteAllBytes(file, bytes);
            var ok = new ReleaseInfo { Version = new Version(9, 0, 0), Size = 2048, Digest = Sha(file) };
            Check("Contrôle du fichier : taille, programme Windows et empreinte corrects acceptés", Updater.Verify(file, ok) && Updater.Verify(file, new ReleaseInfo { Version = ok.Version }));
            Check("Contrôle du fichier : empreinte, taille ou format incorrects refusés", !Updater.Verify(file, new ReleaseInfo { Version = ok.Version, Size = 2048, Digest = "sha256:00" }) && !Updater.Verify(file, new ReleaseInfo { Version = ok.Version, Size = 999 }) && !Updater.Verify(file, new ReleaseInfo { Version = ok.Version, Digest = "md5:abc" }));
            bytes[0] = (byte)'X'; File.WriteAllBytes(file, bytes);
            Check("Contrôle du fichier : fichier qui n'est pas un programme refusé, fichier absent refusé", !Updater.Verify(file, new ReleaseInfo { Version = ok.Version }) && !Updater.Verify(Path.Combine(root, "absent.exe"), ok));
            string self = Assembly.GetExecutingAssembly().Location;
            Check("Contrôle de version : l'exe annonce au moins la version publiée", Updater.AnnouncesVersion(self, new ReleaseInfo { Version = new Version(1, 0, 1) }) && !Updater.AnnouncesVersion(self, new ReleaseInfo { Version = new Version(9, 0, 0) }) && !Updater.AnnouncesVersion(file, new ReleaseInfo { Version = new Version(1, 0, 1) }));
            var found = Updater.CheckAsync(url => ReleaseJson("v9.0.0", good, 2048, null)).GetAwaiter().GetResult();
            Check("Recherche : version plus récente trouvée", found != null && found.Version == new Version(9, 0, 0));
            Check("Recherche : même version ou réponse inutilisable : à jour", Updater.CheckAsync(url => ReleaseJson("v1.0.1", good.Replace("9.0.0", "1.0.1"), 1, null)).GetAwaiter().GetResult() == null && Updater.CheckAsync(url => "n'importe quoi").GetAwaiter().GetResult() == null);
            bool network = false; try { Updater.CheckAsync(url => { throw new WebException("hors ligne"); }).GetAwaiter().GetResult(); } catch (Exception ex) { network = Updater.IsKnownError(ex); }
            Check("Recherche : erreur réseau reconnue pour un message clair", network);
            var downloadable = new ReleaseInfo { Version = new Version(1, 0, 1), AssetUrl = good, Size = new FileInfo(self).Length, Digest = Sha(self) };
            string folder = Path.Combine(root, "telechargement");
            string path = Updater.DownloadForUpdate(downloadable, (url, target) => File.Copy(self, target), folder).GetAwaiter().GetResult();
            Check("Téléchargement : exe vérifié placé dans le dossier de mise à jour", path == Path.Combine(folder, "Korecto.exe") && File.Exists(path) && Sha(path) == downloadable.Digest);
            string failure = null; try { Updater.DownloadForUpdate(downloadable, (url, target) => File.WriteAllBytes(target, new byte[3000]), folder + "-corrompu").GetAwaiter().GetResult(); } catch (Exception ex) { failure = ex.Message; }
            Check("Téléchargement : fichier corrompu refusé", failure != null && failure.Contains("incomplet"));
            failure = null; try { Updater.DownloadForUpdate(new ReleaseInfo { Version = new Version(9, 0, 0), AssetUrl = good, Size = downloadable.Size, Digest = downloadable.Digest }, (url, target) => File.Copy(self, target), folder + "-version-invalide").GetAwaiter().GetResult(); } catch (Exception ex) { failure = ex.Message; }
            Check("Téléchargement : exe dont le numéro ne correspond pas à la release refusé", failure != null && failure.Contains("numéro"));
            Check("Attente : fin détectée, délai dépassé signalé", Updater.WaitUntil(MakeCountdown(3), 2000, 10) && !Updater.WaitUntil(() => false, 150, 10));
            string programs = Path.Combine(root, "Programs"), data = Path.Combine(root, "Data");
            var registry = new FakeInstallRegistry(); var installer = new SelfInstaller(programs, data, Path.Combine(root, "Menu"), registry, f => DeleteFixtureDirectory(f));
            installer.Install(self); Directory.CreateDirectory(data); File.WriteAllText(Path.Combine(data, "ia.json"), "{\"Provider\":\"claude\",\"Model\":\"haiku\"}");
            string newer = Path.Combine(root, "nouveau", "Korecto.exe"); Directory.CreateDirectory(Path.GetDirectoryName(newer)); File.WriteAllBytes(newer, File.ReadAllBytes(self)); File.AppendAllText(newer, "x");
            var polls = MakeCountdown(3);
            Updater.ApplyAndRestart(newer, installer, () => !polls(), 3000);
            Check("Application : attend la fermeture, remplace l'exe installé, réglages conservés", Sha(installer.InstalledExe) == Sha(newer) && File.ReadAllText(Path.Combine(data, "ia.json")).Contains("haiku"));
            failure = null; try { Updater.ApplyAndRestart(newer, installer, () => true, 300); } catch (InvalidOperationException ex) { failure = ex.Message; }
            Check("Application : si Korecto reste ouvert, rien n'est remplacé", failure != null && failure.Contains("fermé"));
            var pref = new UpdatePreference(Path.Combine(root, "prefs"));
            bool initial = pref.Enabled; pref.Save(false); bool off = !pref.Enabled; pref.Save(true);
            Check("Réglage : recherche au démarrage activée par défaut, désactivable puis réactivable", initial && off && pref.Enabled);
            using (var dialog = new SettingsDialog(() => AiSettings.Default(), (o, c) => null, new StartupPreference(new FakeStartupStore()), updatePreference: pref)) {
                Check("Paramètres : version affichée, bouton de recherche inactif sans fonction de recherche", dialog.UpdateNow.Enabled == false && dialog.Updates.Checked);
            }
            using (var dialog = new SettingsDialog(() => AiSettings.Default(), (o, c) => null, new StartupPreference(new FakeStartupStore()), updatePreference: pref, checkUpdates: f => Task.FromResult(true))) {
                Check("Paramètres : bouton de recherche actif avec fonction de recherche", dialog.UpdateNow.Enabled);
            }
        }
        private static Func<bool> MakeCountdown(int polls) { int left = polls; return () => --left <= 0; }
        private static async Task Capture(string directory) {
            bool originalTheme = Palette.IsDark; Palette.Use(true);
            using (var r = new Rig()) {
                r.Window.Opacity = 1;
                r.Window.PerformLayout();
                using (var bitmap = new Bitmap(r.Window.Width, r.Window.Height)) { r.Window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(directory, "ACCUEIL.png")); }
                r.Call("SwitchMode");
                r.Get<TextBox>("manualInput").Text = "Bonjour,\r\n\r\nLes documents sont prêts. Je vous les transmets pour validation.\r\n\r\nBonne journée.";
                r.Window.PerformLayout();
                using (var bitmap = new Bitmap(r.Window.Width, r.Window.Height)) {
                    r.Window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                    bitmap.Save(Path.Combine(directory, "INTERFACE.png"));
                }
                Palette.Use(false); r.Window.PerformLayout();
                using (var bitmap = new Bitmap(r.Window.Width, r.Window.Height)) {
                    r.Window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                    bitmap.Save(Path.Combine(directory, "INTERFACE_CLAIR.png"));
                }
                await r.Ready("Bonjour,\r\n\r\nLes documents sont prêts. Je vous les transmets pour validation.\r\n\r\nBonne journée.");
                r.Runner.Next(); var task = (Task)r.Call("Request", true);
                r.Runner.Pending.SetResult("Bonjour,\r\n\r\nJe vous transmets les documents finalisés afin que vous puissiez les valider.\r\n\r\nBonne journée."); await task;
                using (var bitmap = new Bitmap(r.Window.Width, r.Window.Height)) { r.Window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(directory, "TROIS_VERSIONS.png")); }
            }
            Palette.Use(originalTheme);
            using (var dialog = new SettingsDialog(() => AiSettings.Default(), (owner, current) => null, new StartupPreference(new FakeStartupStore()))) {
                dialog.Show(); dialog.PerformLayout();
                using (var bitmap = new Bitmap(dialog.Width, dialog.Height)) { dialog.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(directory, "PARAMETRES.png")); }
                dialog.Close();
            }
        }
    }
}




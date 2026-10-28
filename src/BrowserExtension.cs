using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CorrecteurMail {
    internal static class BrowserExtension {
        internal static readonly string[] Files = { "manifest.json", "background.js", "status.html", "status.js", "icon-16.png", "icon-48.png", "icon-128.png" };
        internal static readonly string[] RegistryKeys = { @"Software\Google\Chrome\NativeMessagingHosts\com.korecto.selection", @"Software\BraveSoftware\Brave-Browser\NativeMessagingHosts\com.korecto.selection", @"Software\Microsoft\Edge\NativeMessagingHosts\com.korecto.selection", @"Software\Chromium\NativeMessagingHosts\com.korecto.selection" };
        internal static string Prepare(string folder = null, string executable = null, Action<string, string> register = null) {
            folder = Path.GetFullPath(folder ?? Path.Combine(AppFolders.Data, "EXTENSION"));
            executable = Path.GetFullPath(executable ?? Application.ExecutablePath);
            if (!File.Exists(executable)) throw new IOException("Le programme Korecto est introuvable.");
            Directory.CreateDirectory(folder);
            foreach (string name in Files) {
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("CorrecteurMail.Extension." + name)) {
                    if (input == null) throw new IOException("Extension absente de cette copie de Korecto.");
                    using (var output = File.Create(Path.Combine(folder, name))) input.CopyTo(output);
                }
            }
            string manifest = Path.Combine(folder, "com.korecto.selection.json");
            string json = new JavaScriptSerializer().Serialize(new { name = SelectionAccess.Host, description = "Sélection vers Korecto", path = executable, type = "stdio", allowed_origins = new [] { SelectionAccess.ExtensionOrigin } });
            if (register == null) {
                // A development and an installed copy may coexist. Never silently redirect
                // a host belonging to a different Korecto executable.
                foreach (string path in RegistryKeys) using (var key = Registry.CurrentUser.OpenSubKey(path)) {
                    string previous = key == null ? null : key.GetValue("") as string;
                    if (!String.IsNullOrWhiteSpace(previous) && !OwnManifest(previous, executable))
                        throw new IOException("Une autre liaison Korecto existe déjà. Désactivez-la avant de préparer cette copie.");
                }
                register = (keyPath, value) => { using (var key = Registry.CurrentUser.CreateSubKey(keyPath)) key.SetValue("", value, RegistryValueKind.String); };
            }
            File.WriteAllText(manifest, json, new System.Text.UTF8Encoding(false));
            foreach (string key in RegistryKeys) register(key, manifest);
            return folder;
        }
        internal static bool OwnManifest(string file, string executable) {
            try {
                if (!File.Exists(file) || new FileInfo(file).Length > 16384) return false;
                var value = new JavaScriptSerializer().DeserializeObject(File.ReadAllText(file)) as Dictionary<string,object>;
                return value != null && value.ContainsKey("name") && value["name"] as string == SelectionAccess.Host && value.ContainsKey("path") && value["path"] is string && String.Equals(Path.GetFullPath((string)value["path"]), Path.GetFullPath(executable), StringComparison.OrdinalIgnoreCase);
            } catch (Exception ex) { if (!(ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is InvalidOperationException)) throw; return false; }
        }
        internal static string InstallDefault(Func<BrowserTarget> detect = null, Func<string> prepare = null, Action<string> copy = null, Action<BrowserTarget> launch = null) {
            var target = (detect ?? BrowserTarget.Default)();
            if (target == null || !target.Supported) throw new IOException("Votre navigateur par défaut nécessite une version adaptée de l’extension. Alt + K reste disponible.");
            string folder = (prepare ?? (() => Prepare()))();
            (copy ?? (text => Clipboard.SetText(text)))(folder);
            (launch ?? (browser => browser.OpenExtensions()))(target);
            return folder;
        }
        internal static void Show(IWin32Window owner) {
            try { string folder = InstallDefault(); using (var dialog = new BrowserExtensionDialog(folder)) dialog.ShowDialog(owner); }
            catch (Exception ex) {
                if (!(ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is System.ComponentModel.Win32Exception || ex is ArgumentException || ex is System.Runtime.InteropServices.ExternalException)) throw;
                MessageBox.Show(owner,ex.Message,"Extension du navigateur",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }
    }
    internal sealed class BrowserTarget {
        internal string Executable, Name;
        internal bool Supported;
        internal static BrowserTarget Parse(string command) {
            if(String.IsNullOrWhiteSpace(command)) return null;
            var match=System.Text.RegularExpressions.Regex.Match(Environment.ExpandEnvironmentVariables(command), "^\\s*(?:\"([^\"]+\\.exe)\"|(.+?\\.exe)(?:\\s|$))", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if(!match.Success) return null;
            string exe=match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
            string name=Path.GetFileNameWithoutExtension(exe).ToLowerInvariant();
            var names=new Dictionary<string,string> { {"chrome","Chrome"},{"brave","Brave"},{"msedge","Microsoft Edge"},{"vivaldi","Vivaldi"},{"chromium","Chromium"} };
            return new BrowserTarget { Executable=exe, Name=names.ContainsKey(name) ? names[name] : name, Supported=names.ContainsKey(name) };
        }
        internal static BrowserTarget Default() {
            foreach(string protocol in new [] {"https","http"}) {
                using(var choice=Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\"+protocol+@"\UserChoice")) {
                    string id=choice==null ? null : choice.GetValue("ProgId") as string;
                    if(String.IsNullOrWhiteSpace(id) || id.IndexOfAny(new[]{'\\','/'})>=0) continue;
                    using(var command=Registry.ClassesRoot.OpenSubKey(id+@"\shell\open\command")) {
                        var target=Parse(command==null ? null : command.GetValue("") as string);
                        if(target!=null) return target;
                    }
                }
            }
            return null;
        }
        internal void OpenExtensions() {
            if(!Supported || !File.Exists(Executable)) throw new IOException("Le navigateur par défaut est introuvable ou incompatible avec cette version de test.");
            string url=Name=="Microsoft Edge" ? "edge://extensions" : Name=="Brave" ? "brave://extensions" : Name=="Vivaldi" ? "vivaldi://extensions" : "chrome://extensions";
            Process.Start(new ProcessStartInfo(Executable,url) { UseShellExecute=true });
        }
    }
    internal sealed class BrowserExtensionDialog : Form {
        internal readonly TextBox Folder = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        internal BrowserExtensionDialog(string folder = "") {
            Text = "Extension du navigateur · Korecto"; Font = new Font("Segoe UI",9.5f);
            AutoScaleDimensions = new SizeF(96,96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(490,260); FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false; StartPosition = FormStartPosition.CenterParent; AppBrand.SetIcon(this);
            var layout=new TableLayoutPanel { Dock=DockStyle.Fill, Padding=new Padding(20), ColumnCount=1, RowCount=5 };
            foreach(int height in new[]{38,44,78,28,32}) layout.RowStyles.Add(new RowStyle(SizeType.Absolute,height));
            layout.Controls.Add(new Label { Text="Terminer dans votre navigateur", Font=new Font("Segoe UI Semibold",15), Dock=DockStyle.Fill },0,0);
            layout.Controls.Add(new Label { Text="La page des extensions est ouverte. Le chemin du dossier est déjà copié.", Dock=DockStyle.Fill },0,1);
            layout.Controls.Add(new Label { Text="1. Activez Mode développeur.\r\n2. Cliquez sur Charger l’extension non empaquetée.\r\n3. Collez le chemin avec Ctrl + V, puis validez.\r\nCette installation se fait une seule fois.", Dock=DockStyle.Fill },0,2);
            Folder.Text=folder; Folder.Enter+=(s,e)=>Folder.SelectAll(); layout.Controls.Add(Folder,0,3);
            var close=new SoftButton { Text="Fermer", DialogResult=DialogResult.Cancel, Dock=DockStyle.Right, Width=100 };
            layout.Controls.Add(close,0,4); CancelButton=close; Controls.Add(layout); Palette.Apply(this); HandleCreated+=(s,e)=>Palette.ApplyCaption(this);
        }
    }
}
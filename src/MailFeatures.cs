using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace CorrecteurMail {
    // Explicitly entered technical terms only; no harvesting from messages or address books.
    internal sealed class VocabularyStore {
        private readonly string path;
        internal VocabularyStore(string file = null) { path = file ?? Path.Combine(AppFolders.Data, "lexique-metier.txt"); }
        internal static string[] Parse(string text) {
            if ((text ?? "").Length > 8000) throw new InvalidOperationException("Dictionnaire trop long (8 000 caractères maximum).");
            var terms = new System.Collections.Generic.List<string>();
            foreach (string line in (text ?? "").Replace("\r", "").Split('\n')) {
                string term = line.Trim(); if (term.Length == 0) continue;
                if (term.Length > 80 || Regex.IsMatch(term, @"[\p{Cc}\p{Cf}]")) throw new InvalidOperationException("Un terme doit tenir sur une ligne de 80 caractères maximum, sans caractère de contrôle.");
                if (!terms.Contains(term)) terms.Add(term);
            }
            return terms.ToArray();
        }
        internal string[] Load() {
            if (!File.Exists(path)) return new string[0];
            if (new FileInfo(path).Length > 32000) throw new InvalidOperationException("Dictionnaire métier trop volumineux.");
            return Parse(File.ReadAllText(path, Encoding.UTF8));
        }
        internal void Save(string text) {
            string normalized = String.Join("\r\n", Parse(text));
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            string temp = path + ".tmp";
            File.WriteAllText(temp, normalized, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
        }
    }
    internal sealed class VocabularyDialog : Form {
        internal readonly TextBox Terms = new TextBox();
        internal readonly Button Save = new SoftButton { Primary = true, Text = "Enregistrer" };
        internal VocabularyDialog(VocabularyStore store = null) {
            SuspendLayout(); AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            store = store ?? new VocabularyStore();
            Text = "Dictionnaire métier · Korecto"; Font = new Font("Segoe UI", 10);
            ClientSize = new Size(490, 380); MinimumSize = Size;
            StartPosition = FormStartPosition.CenterParent; MinimizeBox = MaximizeBox = false; AppBrand.SetIcon(this);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18), RowCount = 4, ColumnCount = 1 };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            panel.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Termes techniques, sigles et matériaux à préserver, un par ligne. Sans noms de clients. Les termes sont transmis avec le texte lors de vos demandes à l’IA." }, 0, 0);
            Terms.Multiline = true; Terms.AcceptsReturn = true; Terms.ScrollBars = ScrollBars.Vertical; Terms.Dock = DockStyle.Fill; Terms.MaxLength = 8000;
            var message = new Label { Dock = DockStyle.Fill, ForeColor = Palette.Muted };
            try { Terms.Text = String.Join("\r\n", store.Load()); }
            catch (Exception ex) { if (!(ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException)) throw; message.Text = "Lecture impossible : " + ex.Message; Save.Enabled = false; }
            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 }; buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var cancel = new SoftButton { Text = "Annuler", DialogResult = DialogResult.Cancel, Dock = DockStyle.Fill };
            Save.Dock = DockStyle.Fill;
            Save.Click += (s,e) => { try { store.Save(Terms.Text); DialogResult = DialogResult.OK; } catch (Exception ex) { if (!(ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException)) throw; message.Text = ex.Message; } };
            buttons.Controls.Add(cancel); buttons.Controls.Add(Save); panel.Controls.Add(Terms, 0, 1); panel.Controls.Add(message, 0, 2); panel.Controls.Add(buttons, 0, 3); Controls.Add(panel);
            CancelButton = cancel; Palette.Apply(this); AutoScaleDimensions = new SizeF(96, 96); ResumeLayout(true);
        }
    }
    internal static class SubjectText {
        internal static string Validate(string text, string source) {
            string subject = (text ?? "").Trim();
            if (subject.Length == 0 || subject.Length > 180 || Regex.IsMatch(subject, @"[\p{Cc}\p{Cf}\u2028\u2029]")) throw new InvalidOperationException("Objet invalide : une seule ligne de 180 caractères maximum est attendue.");
            var numbers = new System.Collections.Generic.HashSet<string>();
            foreach (Match number in Regex.Matches(source, @"\d+(?:[.,]\d+)*")) numbers.Add(number.Value);
            foreach (Match number in Regex.Matches(subject, @"\d+(?:[.,]\d+)*")) if (!numbers.Contains(number.Value)) throw new InvalidOperationException("L’objet contient un chiffre absent du texte. Demandez une nouvelle proposition.");
            return subject;
        }
    }
    internal sealed class SubjectDialog : Form {
        internal readonly TextBox Proposal = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        internal readonly Button Generate = new SoftButton { Text = "Proposer un objet", Primary = true, Dock = DockStyle.Fill };
        internal readonly Button Copy = new SoftButton { Text = "Copier l’objet", Dock = DockStyle.Fill, Enabled = false };
        private bool disposed;
        internal SubjectDialog(Func<Task<string>> generate, Action<string> copy = null) {
            SuspendLayout(); AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            Text = "Objet du mail · Korecto"; Font = new Font("Segoe UI", 10); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(520, 220); MinimumSize = Size; StartPosition = FormStartPosition.CenterParent; MinimizeBox = MaximizeBox = false; AppBrand.SetIcon(this);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18), ColumnCount = 1, RowCount = 4 };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 48)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            panel.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Proposition à partir du texte affiché dans Korecto. Copiez l’objet pour le coller dans votre mail." }, 0, 0);
            var message = new Label { Dock = DockStyle.Fill, ForeColor = Palette.Muted, Text = "Cliquez pour envoyer le texte à l’IA." };
            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 }; buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); buttons.Controls.Add(Generate); buttons.Controls.Add(Copy);
            Generate.Click += async (s,e) => {
                Generate.Enabled = Copy.Enabled = false; Proposal.Clear(); message.Text = "Proposition d’objet en cours…";
                try { string text = await generate(); if (disposed) return; Proposal.Text = text; Copy.Enabled = true; message.Text = "Vérifiez l’objet avant de le copier."; }
                catch (Exception ex) { if (!(ex is InvalidOperationException || ex is System.IO.IOException)) throw; if (!disposed) message.Text = ex is InvalidOperationException ? ex.Message : "Objet indisponible. Réessayez."; }
                finally { if (!disposed) Generate.Enabled = true; }
            };
            Copy.Click += (s,e) => { try { if (copy != null) copy(Proposal.Text); else Clipboard.SetText(Proposal.Text); message.Text = "Objet copié."; } catch (System.Runtime.InteropServices.ExternalException) { message.Text = "Presse-papiers occupé. Réessayez."; } };
            FormClosed += (s,e) => disposed = true;
            Disposed += (s,e) => disposed = true;
            panel.Controls.Add(Proposal, 0, 1); panel.Controls.Add(message, 0, 2); panel.Controls.Add(buttons, 0, 3); Controls.Add(panel); Palette.Apply(this); AutoScaleDimensions = new SizeF(96, 96); ResumeLayout(true);
        }
    }
}

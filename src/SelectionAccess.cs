using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CorrecteurMail {
    internal sealed class InternalStartupStore : IStartupStore { public string Read() { return null; } public void Write(string command) { throw new IOException("Démarrage Windows désactivé pour cette version de test."); } public void Remove() { } }
    internal sealed class SelectionAccess : NativeWindow, IDisposable {
        internal const int Limit = 30000;
        private const int Hotkey = 0x4B01;
        private readonly Action<string> receive, report;
        private bool capturing, registered;
        [DllImport("user32.dll", SetLastError=true)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr window, int id);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
        internal SelectionAccess(Action<string> receive, Action<string> report) {
            this.receive=receive; this.report=report;
            CreateHandle(new CreateParams());
            registered=RegisterHotKey(Handle, Hotkey, 0x4001, 0x4B); // Alt+K, no repeat
            if (!registered) report("Alt + K indisponible : un autre programme utilise ce raccourci.");
        }
        protected override void WndProc(ref Message message) {
            if (message.Msg==0x0312 && message.WParam.ToInt32()==Hotkey) { Capture(); message.Result=IntPtr.Zero; return; }
            base.WndProc(ref message);
        }
        internal static bool Valid(string text) { return !String.IsNullOrWhiteSpace(text) && text.Length<=Limit && text.IndexOf('\0')<0; }
        private async void Capture() {
            if(capturing) return; capturing=true;
            try {
                IntPtr source=Native.GetForegroundWindow();
                uint process; Native.GetWindowThreadProcessId(source,out process);
                if(process==(uint)Process.GetCurrentProcess().Id) { report("Sélectionnez d’abord un texte dans une autre application, puis Alt + K."); return; }
                var watch=Stopwatch.StartNew();
                while ((GetAsyncKeyState(0x12)<0 || GetAsyncKeyState(0x4B)<0 || GetAsyncKeyState(0x11)<0 || GetAsyncKeyState(0x10)<0) && watch.ElapsedMilliseconds<1500) await Task.Delay(15);
                if(GetAsyncKeyState(0x12)<0 || GetAsyncKeyState(0x11)<0 || GetAsyncKeyState(0x10)<0 || Native.GetForegroundWindow()!=source) return;
                // Materialize the previous formats before the source application replaces them.
                var previous=new DataObject(); var old=Clipboard.GetDataObject();
                if(old!=null) foreach(string format in old.GetFormats(false)) { object value=old.GetData(format,false); if(value!=null) previous.SetData(format,false,value); }
                uint before=GetClipboardSequenceNumber();
                keybd_event(0x11,0,0,UIntPtr.Zero); keybd_event(0x43,0,0,UIntPtr.Zero); keybd_event(0x43,0,2,UIntPtr.Zero); keybd_event(0x11,0,2,UIntPtr.Zero);
                watch.Restart();
                while(GetClipboardSequenceNumber()==before && watch.ElapsedMilliseconds<1200) await Task.Delay(20);
                uint copied=GetClipboardSequenceNumber();
                if(copied==before) { report("Aucune sélection reçue. Sélectionnez du texte, puis Alt + K."); return; }
                string text=null;
                try { if(Native.GetForegroundWindow()==source && Clipboard.ContainsText()) text=Clipboard.GetText(); }
                finally { if(GetClipboardSequenceNumber()==copied) { if(old==null) Clipboard.Clear(); else Clipboard.SetDataObject(previous,true); } }
                if(Valid(text)) receive(text); else report("Sélection vide ou trop longue (30 000 caractères maximum).");
            } catch(Exception ex) { if(!(ex is ExternalException || ex is System.Threading.ThreadStateException || ex is ArgumentException)) throw; report("Sélection inaccessible : réessayez ou copiez-collez le texte en mode Manuel."); }
            finally { capturing=false; }
        }
        public void Dispose() { if(registered) UnregisterHotKey(Handle,Hotkey); registered=false; DestroyHandle(); }
    }
}

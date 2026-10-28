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
        internal const string Host = "com.korecto.selection";
        internal const string ExtensionOrigin = "chrome-extension://gfdaclbaflbekblfmkcldcpeddmfnakm/";
        private const int Hotkey = 0x4B01, CopyData = 0x004A;
        private readonly Action<string> receive, report;
        private bool capturing, registered;
        [StructLayout(LayoutKind.Sequential)] private struct CopyPacket { public IntPtr Tag; public int Length; public IntPtr Data; }
        [DllImport("user32.dll", SetLastError=true)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr window, int id);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
        [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr FindWindow(string name, string caption);
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr sender, ref CopyPacket packet, uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll")] private static extern bool AllowSetForegroundWindow(uint process);
        internal static string ReceiverName(bool internalMode) { return "Korecto.Selection." + Process.GetCurrentProcess().SessionId + "." + Environment.UserName + (internalMode ? ".internal" : ".installed"); }
        internal SelectionAccess(bool internalMode, Action<string> receive, Action<string> report, bool enableHotkey = true) {
            this.receive=receive; this.report=report;
            CreateHandle(new CreateParams { Caption=ReceiverName(internalMode) });
            registered=enableHotkey && RegisterHotKey(Handle, Hotkey, 0x4001, 0x4B); // Alt+K, no repeat
            if (enableHotkey && !registered) report("Alt + K indisponible : un autre programme utilise ce raccourci. Le menu du navigateur reste utilisable.");
        }
        protected override void WndProc(ref Message message) {
            if (message.Msg==0x0312 && message.WParam.ToInt32()==Hotkey) { Capture(); message.Result=IntPtr.Zero; return; }
            if (message.Msg==CopyData) {
                try {
                    var packet=(CopyPacket)Marshal.PtrToStructure(message.LParam, typeof(CopyPacket));
                    if(packet.Tag.ToInt64()!=Hotkey || packet.Length<2 || packet.Length>(Limit+1)*2 || packet.Length%2!=0 || packet.Data==IntPtr.Zero) { message.Result=IntPtr.Zero; return; }
                    string text=Marshal.PtrToStringUni(packet.Data, packet.Length/2);
                    if(!Valid(text)) { message.Result=IntPtr.Zero; return; }
                    receive(text); message.Result=new IntPtr(1); return;
                } catch (ArgumentException) { message.Result=IntPtr.Zero; return; }
            }
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
                if(copied==before) { report("Aucune sélection reçue. Sélectionnez du texte ou utilisez le menu Korecto du navigateur."); return; }
                string text=null;
                try { if(Native.GetForegroundWindow()==source && Clipboard.ContainsText()) text=Clipboard.GetText(); }
                finally { if(GetClipboardSequenceNumber()==copied) { if(old==null) Clipboard.Clear(); else Clipboard.SetDataObject(previous,true); } }
                if(Valid(text)) receive(text); else report("Sélection vide ou trop longue (30 000 caractères maximum).");
            } catch(Exception ex) { if(!(ex is ExternalException || ex is System.Threading.ThreadStateException || ex is ArgumentException)) throw; report("Sélection inaccessible : réessayez ou utilisez le menu Korecto du navigateur."); }
            finally { capturing=false; }
        }
        internal static bool Send(string text, bool internalMode) {
            if(!Valid(text)) return false;
            IntPtr window=FindWindow(null,ReceiverName(internalMode)); if(window==IntPtr.Zero) return false;
            uint process; Native.GetWindowThreadProcessId(window,out process); AllowSetForegroundWindow(process);
            IntPtr data=Marshal.StringToHGlobalUni(text);
            try { var packet=new CopyPacket { Tag=new IntPtr(Hotkey), Length=text.Length*2, Data=data }; IntPtr result; return SendMessageTimeout(window,CopyData,IntPtr.Zero,ref packet,0x0002,2000,out result)!=IntPtr.Zero && result.ToInt64()==1; }
            finally { Marshal.FreeHGlobal(data); }
        }
        internal static byte[] ReadFrame(Stream input) {
            var header=new byte[4]; ReadExact(input,header); int length=BitConverter.ToInt32(header,0);
            if(length<=0 || length>200000) throw new InvalidDataException("Message trop volumineux.");
            var body=new byte[length]; ReadExact(input,body); return body;
        }
        private static void ReadExact(Stream stream, byte[] bytes) { int offset=0; while(offset<bytes.Length) { int count=stream.Read(bytes,offset,bytes.Length-offset); if(count==0) throw new EndOfStreamException(); offset+=count; } }
        internal static void WriteFrame(Stream output, object response) { byte[] body=Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(response)); output.Write(BitConverter.GetBytes(body.Length),0,4); output.Write(body,0,body.Length); output.Flush(); }
        internal static string ParseSelection(byte[] frame) {
            var message=new JavaScriptSerializer().DeserializeObject(new UTF8Encoding(false,true).GetString(frame)) as Dictionary<string,object>;
            if(message==null || !message.ContainsKey("action") || message["action"] as string!="selection" || !message.ContainsKey("text") || !(message["text"] is string) || !Valid((string)message["text"])) throw new InvalidDataException("Sélection invalide.");
            return (string)message["text"];
        }
        internal static void BrowserHost(string origin) {
            using(var output=Console.OpenStandardOutput()) {
                try {
                    if(origin!=ExtensionOrigin) throw new InvalidDataException("Extension non autorisée.");
                    string text; using(var input=Console.OpenStandardInput()) text=ParseSelection(ReadFrame(input));
                    bool internalMode=!new SelfInstaller().IsInstalledCopy(Application.ExecutablePath);
                    if(!Send(text,internalMode)) {
                        Process.Start(new ProcessStartInfo(Application.ExecutablePath,internalMode ? "--internal --selection-start" : "--selection-start") { UseShellExecute=false });
                        bool delivered=false;
                        for(int i=0;i<40;i++) { Thread.Sleep(200); if(Send(text,internalMode)) { delivered=true; break; } }
                        if(!delivered) throw new InvalidOperationException("Ouvrez Korecto et réessayez la sélection.");
                    }
                    WriteFrame(output,new { ok=true });
                } catch(Exception ex) {
                    if(!(ex is IOException || ex is InvalidDataException || ex is ArgumentException || ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)) throw;
                    WriteFrame(output,new { ok=false, error=ex.Message });
                }
            }
        }
        public void Dispose() { if(registered) UnregisterHotKey(Handle,Hotkey); registered=false; DestroyHandle(); }
    }
}

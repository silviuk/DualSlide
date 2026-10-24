using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
public class BusyClosingDeck {
 readonly dynamic actual;bool first=true;public BusyClosingDeck(object deck){actual=deck;}
 public int Saved{set{if(first){first=false;throw new COMException("Injected busy close",unchecked((int)0x80010001));}actual.Saved=value;}}
 public void Close(){actual.Close();}
}
class ShutdownTest {
 static object Field(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Set(MainForm f,string n,object value){typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(f,value);}
 static void Call(MainForm f,string n){typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 static void Assert(bool condition,string message){if(!condition)throw new Exception(message);}
 static string Hash(string path){using(var sha=SHA256.Create())using(var file=File.OpenRead(path))return Convert.ToBase64String(sha.ComputeHash(file));}
 static void Pump(){Application.DoEvents();Thread.Sleep(20);}
 static void WaitStopped(MainForm f){var end=DateTime.UtcNow.AddSeconds(5);while(((bool)Field(f,"running")||(bool)Field(f,"stopping"))&&DateTime.UtcNow<end)Pump();Assert(!(bool)Field(f,"running")&&!(bool)Field(f,"stopping"),"Pair did not finish closing");}
 static void Locked(string file){bool denied=false;try{using(var stream=new FileStream(file,FileMode.Open,FileAccess.Write,FileShare.ReadWrite)) {}}catch(IOException){denied=true;}Assert(denied,"Original file was not write-protected during show");}
 [STAThread]static void Main(){Native.SetProcessDpiAwarenessContext(new IntPtr(-4));AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/shutdown-settings"));string a=Path.GetFullPath("test-master.pptx"),b=Path.GetFullPath("test-follower.pptx");string ah=Hash(a),bh=Hash(b);DateTime at=File.GetLastWriteTimeUtc(a),bt=File.GetLastWriteTimeUtc(b);dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")),unrelated=null;
 try{unrelated=app.Presentations.Add(-1);unrelated.Slides.Add(1,12);int baseline=app.Presentations.Count;using(var form=new MainForm()){form.Show();((ComboBox)Field(form,"masterPath")).Text=a;((ComboBox)Field(form,"followerPath")).Text=b;
 for(int mode=0;mode<5;mode++){Call(form,"Start");Assert((bool)Field(form,"running"),"Start failed");dynamic master=Field(form,"master"),follower=Field(form,"follower");IntPtr mh=(IntPtr)Field(form,"masterHandle"),fh=(IntPtr)Field(form,"followerHandle");Assert((int)master.ReadOnly==-1&&(int)follower.ReadOnly==-1,"Decks not read-only");Assert((string)master.FullName==a&&(string)follower.FullName==b,"Opened a copy instead of original path");Locked(a);Locked(b);
 master.Slides[1].Shapes.AddTextbox(1,20,20,200,30).TextFrame.TextRange.Text="Discard this in-memory test edit";follower.Slides[1].Shapes.AddTextbox(1,20,20,200,30).TextFrame.TextRange.Text="Discard this in-memory test edit";master.Saved=0;follower.Saved=0;
 if(mode==0){dynamic show=Field(form,"masterShow");show.View.Exit();Call(form,"CheckSession");}
 else if(mode==1){dynamic show=Field(form,"followerShow");show.View.Exit();Call(form,"CheckSession");}
 else if(mode==2){Set(form,"paused",true);dynamic show=Field(form,"masterShow");show.View.Exit();Call(form,"CheckSession");}
 else if(mode==3){Set(form,"master",new BusyClosingDeck((object)master));Call(form,"Stop");Assert((bool)Field(form,"stopping"),"Busy close lost its pending cleanup");Locked(a);Locked(b);Call(form,"CheckSession");}
 else{Native.SetForegroundWindow(mh);IntPtr memory=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Native.Key)));try{Marshal.StructureToPtr(new Native.Key{Code=27,Scan=1},memory,false);var result=(IntPtr)typeof(MainForm).GetMethod("Keyboard",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{0,new IntPtr(0x100),memory});Assert(result==new IntPtr(1),"Escape was allowed to reach PowerPoint");}finally{Marshal.FreeHGlobal(memory);}}
 WaitStopped(form);Assert(!Native.IsWindow(mh)&&!Native.IsWindow(fh),"A slideshow remains open");Assert((int)app.Presentations.Count==baseline,"Owned documents left open or unrelated presentation closed");Assert((IntPtr)Field(form,"keyHook")==IntPtr.Zero&&(IntPtr)Field(form,"mouseHook")==IntPtr.Zero,"Input hooks retained");Assert(Hash(a)==ah&&Hash(b)==bh&&File.GetLastWriteTimeUtc(a)==at&&File.GetLastWriteTimeUtc(b)==bt,"Original bytes or timestamps changed");using(var stream=new FileStream(a,FileMode.Open,FileAccess.Write,FileShare.ReadWrite)){}Console.WriteLine("PASS: shutdown path "+mode+", both shows/documents closed without saving; originals locked during session and unchanged afterward.");}
 form.Close();}Console.WriteLine("PASS: original paths, read-only mode, source hashes/timestamps, unrelated document, write-lock release, and busy-close retry.");}
 catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}finally{try{unrelated.Saved=-1;unrelated.Close();}catch{}}}
}

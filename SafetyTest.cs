using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
class SafetyTest {
 static object Field(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
 static void Assert(bool value,string text){if(!value)throw new Exception(text);}
 [STAThread]static void Main(string[] args){
  if(args.Length>0){using(var gate=new SingleInstance()){Environment.ExitCode=gate.IsFirst?7:0;}return;}
  try{
   using(var gate=new SingleInstance()){
    Assert(gate.IsFirst,"Unexpected controller running");
    var process=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"child"){UseShellExecute=false,CreateNoWindow=true});
    Assert(process.WaitForExit(4000)&&process.ExitCode==0,"Second process acquired singleton");process.Dispose();
   }
   using(var gate=new SingleInstance())Assert(gate.IsFirst,"Singleton not released");
   Console.WriteLine("PASS: competing process excluded; singleton released for next launch.");
   Native.SetProcessDpiAwarenessContext(new IntPtr(-4));AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/safety-settings"));
   using(var f=new MainForm()){
    f.Show();((ComboBox)Field(f,"masterPath")).Text=Path.GetFullPath("test-master.pptx");((ComboBox)Field(f,"followerPath")).Text=Path.GetFullPath("test-follower.pptx");Call(f,"Start");
    var until=DateTime.UtcNow.AddSeconds(1);while(DateTime.UtcNow<until){Application.DoEvents();Thread.Sleep(20);}
    IntPtr mh=(IntPtr)Field(f,"masterHandle"),fh=(IntPtr)Field(f,"followerHandle");
    ((SessionWatchdog)Field(f,"watchdog")).Dispose();bool recovered=false;
    var watchdog=new SessionWatchdog(()=>{Call(f,"RecoverDesktop");recovered=true;},1000);
    typeof(MainForm).GetField("watchdog",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(f,watchdog);
    watchdog.Track(mh,fh);Thread.Sleep(2200); // deliberately no UI message pumping
    Assert(recovered,"Watchdog was blocked by controller UI");
    Assert(IsIconic(mh)&&IsIconic(fh),"Watchdog failed to release fullscreen windows");
    var pump=(InputPump)Field(f,"input");
    Console.WriteLine("PASS: watchdog recovered desktop while controller thread was blocked.");
    until=DateTime.UtcNow.AddSeconds(8);while((Native.IsWindow(mh)||Native.IsWindow(fh)||(bool)Field(f,"running")||(bool)Field(f,"stopping"))&&DateTime.UtcNow<until){Application.DoEvents();Thread.Sleep(20);}
    Assert(!(bool)Field(f,"running")&&!(bool)Field(f,"stopping"),"Recovery cleanup failed");Assert(!Native.IsWindow(mh)&&!Native.IsWindow(fh),"Shows remain after recovery");
    Assert((IntPtr)Field(f,"keyHook")==IntPtr.Zero&&(IntPtr)Field(f,"mouseHook")==IntPtr.Zero,"Hooks retained after recovery");
    Console.WriteLine("PASS: recovery closed both owned decks and removed input hooks.");f.Close();
   }
  }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}
 }
}

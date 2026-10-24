using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
class PowerPointLifecycleTest {
 static void Set(MainForm f,string n,object value){typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(f,value);}
 static object Get(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){try{typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}catch(TargetInvocationException ex){throw ex.InnerException;}}
 static bool OfficeRunning(){var p=Process.GetProcessesByName("POWERPNT");try{return p.Length>0;}finally{foreach(var x in p)x.Dispose();}}
 [STAThread]static void Main(string[] args){try{
  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));
  if(args.Length>0&&args[0]=="prepare"){
   dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
   foreach(string name in new[]{"test-master.pptx","test-follower.pptx"}){dynamic deck=app.Presentations.Add(-1);deck.Slides.Add(1,12);deck.SaveAs(Path.Combine(args[1],name));deck.Close();}app.Quit();return;
  }
  if(OfficeRunning())throw new Exception("Run this disposable integration test with PowerPoint closed.");
  string folder=Path.GetFullPath(".build/lifecycle-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
  using(var child=Process.Start(new ProcessStartInfo(Application.ExecutablePath,"prepare \""+folder+"\""){UseShellExecute=false})){child.WaitForExit();if(child.ExitCode!=0)throw new Exception("Deck generation failed");}
  WaitForOfficeExit();
  AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/lifecycle-settings"));
  using(var f=new MainForm()){
   f.Show();((ComboBox)Get(f,"masterPath")).Text=Path.Combine(folder,"test-master.pptx");((ComboBox)Get(f,"followerPath")).Text=Path.Combine(folder,"test-follower.pptx");
   Call(f,"Start");if(!(bool)Get(f,"running")||!(bool)Get(f,"ownPowerPoint"))throw new Exception("Owned session failed to start");
   var end=DateTime.UtcNow.AddSeconds(2);while(DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(20);}
   Call(f,"Stop");end=DateTime.UtcNow.AddSeconds(10);while((bool)Get(f,"stopping")&&DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(20);}
   if((bool)Get(f,"stopping"))throw new Exception("Stop did not complete");f.Close();
  }
  WaitForOfficeExit();Console.WriteLine("PASS: DualSlide started PowerPoint, stopped both shows, and PowerPoint process exited while controller remained alive.");
 }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}}
 static void WaitForOfficeExit(){var end=DateTime.UtcNow.AddSeconds(10);while(OfficeRunning()&&DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(100);}if(OfficeRunning())throw new Exception("PowerPoint process remains after cleanup");}
}

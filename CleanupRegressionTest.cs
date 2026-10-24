using System;
using System.Reflection;
using System.Runtime.InteropServices;
public sealed class CleanupPresentations {public int Count;}
public sealed class CleanupApplication {
 public CleanupPresentations Presentations=new CleanupPresentations();public int Quits;public bool Busy;
 public void Quit(){if(Busy){Busy=false;throw new COMException("Busy",unchecked((int)0x80010001));}Quits++;}
}
class CleanupRegressionTest {
 static void Set(MainForm f,string n,object value){typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(f,value);}
 static bool Close(MainForm f){return (bool)typeof(MainForm).GetMethod("ClosePowerPoint",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 static void Assert(bool value,string message){if(!value)throw new Exception(message);}
 [STAThread]static void Main(){AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",System.IO.Path.GetFullPath(".build/cleanup-settings"));
  using(var f=new MainForm())try{
   var app=new CleanupApplication();Set(f,"ppt",app);Set(f,"ownPowerPoint",true);Assert(Close(f)&&app.Quits==1,"Owned empty application did not quit");
   app=new CleanupApplication();Set(f,"ppt",app);Set(f,"ownPowerPoint",false);Assert(Close(f)&&app.Quits==0,"Borrowed application was quit");
   app=new CleanupApplication();app.Presentations.Count=1;Set(f,"ppt",app);Set(f,"ownPowerPoint",true);Assert(Close(f)&&app.Quits==0,"User's new presentation was closed");
   app=new CleanupApplication{Busy=true};Set(f,"ppt",app);Set(f,"ownPowerPoint",true);Assert(!Close(f)&&app.Quits==0,"Busy Quit lost its retry");Assert(Close(f)&&app.Quits==1,"Quit retry failed");Assert(Close(f)&&app.Quits==1,"Quit was repeated after cleanup");
   Console.WriteLine("PASS: owned empty PowerPoint quit, borrowed application/new documents preserved, busy Quit retried, repeated cleanup safe.");
  }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}finally{((PointerForm)typeof(MainForm).GetField("pointer",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f)).Dispose();}
 }
}

using System;

using System.Reflection;

using System.Threading;

using System.Windows.Forms;

public class BusyShow { public BusyView View=new BusyView(); }

public class BusyView { public bool LaserPointerEnabled {get{return false;}} public int State {get{throw new System.Runtime.InteropServices.COMException("Injected busy transition",unchecked((int)0x80010001));}} }

class SessionTest {

 static object Field(MainForm form,string name){return typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);}

 static void Call(MainForm form,string name,params object[] args){typeof(MainForm).GetMethod(name,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,args);}

 static void Pump(){for(int i=0;i<20;i++){Application.DoEvents();Thread.Sleep(25);}}

 [STAThread] static void Main(string[] args){

  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",System.IO.Path.GetFullPath(".build/test-settings"));

  Console.WriteLine("Displays detected: "+Screen.AllScreens.Length);

  foreach(var screen in Screen.AllScreens)Console.WriteLine("Display: "+screen.DeviceName+" "+screen.Bounds);

  if(Screen.AllScreens.Length<2){Console.WriteLine("Two displays required");Environment.ExitCode=2;return;}

  using(var form=new MainForm()){

   try{

    form.Show();

    ((ComboBox)Field(form,"masterPath")).Text=System.IO.Path.GetFullPath("test-master.pptx");

    ((ComboBox)Field(form,"followerPath")).Text=System.IO.Path.GetFullPath("test-follower.pptx");

    ((ComboBox)Field(form,"masterMonitor")).SelectedIndex=args.Length>0?1:0;Console.WriteLine("Calling Start");Call(form,"Start");Console.WriteLine("Start returned; pumping UI");Pump();Console.WriteLine("Initial pump returned");

    if(!(bool)Field(form,"running"))throw new Exception("Session failed to start");

    dynamic master=Field(form,"masterShow"),follower=Field(form,"followerShow");

    Console.WriteLine("Master: "+Native.Client((IntPtr)Field(form,"masterHandle")));

    Console.WriteLine("Follower: "+Native.Client((IntPtr)Field(form,"followerHandle")));

    if(Native.Client((IntPtr)Field(form,"masterHandle"))!=Screen.AllScreens[args.Length>0?1:0].Bounds||Native.Client((IntPtr)Field(form,"followerHandle"))!=Screen.AllScreens[args.Length>0?0:1].Bounds)throw new Exception("Placement mismatch");

    Console.WriteLine("Canvases: "+Native.Client((IntPtr)Field(form,"masterCanvas"))+" / "+Native.Client((IntPtr)Field(form,"followerCanvas")));
    if(Native.Client((IntPtr)Field(form,"masterCanvas"))!=Screen.AllScreens[args.Length>0?1:0].Bounds||Native.Client((IntPtr)Field(form,"followerCanvas"))!=Screen.AllScreens[args.Length>0?0:1].Bounds)throw new Exception("Slideshow canvas placement mismatch");
    Console.WriteLine("PASS: outer frames and actual native slide canvases fill selected displays.");
    Console.WriteLine("Foreground master: "+CallActive(form));

    for(int i=0;i<12;i++){

     if(i%4==0){master.View.GotoSlide(1,-1);follower.View.GotoSlide(1,-1);Pump();}

     master.View.Next();Call(form,"Command",39);Pump();

     int a=(int)master.View.Slide.SlideIndex,b=(int)follower.View.Slide.SlideIndex;

     int ca=(int)master.View.GetClickIndex(),cb=(int)follower.View.GetClickIndex();

     Console.WriteLine(i+": master="+a+"/"+ca+" follower="+b+"/"+cb);

     if(a!=b||ca!=cb)throw new Exception("Slide/build mismatch");

    }

    master.View.State=3;Call(form,"MatchNativeState");Pump();

    if((int)follower.View.State!=3)throw new Exception("Blackout mismatch");

    master.View.State=4;Call(form,"MatchNativeState");Pump();

    if((int)follower.View.State!=4)throw new Exception("Whiteout mismatch");

    master.View.State=1;master.View.GotoSlide(3,-1);Call(form,"MatchNativeState");Pump();

    master.View.Previous();Call(form,"MatchNativeState");Pump();

    if((int)master.View.Slide.SlideIndex!=(int)follower.View.Slide.SlideIndex)throw new Exception("Backward navigation mismatch");

    Native.SetForegroundWindow((IntPtr)Field(form,"masterHandle"));Pump();Call(form,"CheckSession");

    var pointer=(PointerForm)Field(form,"pointer");long ticks=pointer.FrameTicks;System.Threading.Thread.Sleep(250);if(pointer.FrameTicks-ticks<5)throw new Exception("Pointer thread stalled");pointer.Hide();System.Threading.Thread.Sleep(40);if(pointer.Visible)throw new Exception("Pointer did not hide");

    Console.WriteLine("PASS: independent pointer thread and hide.");

    Console.WriteLine("PASS: 12 paired native advances, blackout, whiteout, and backward navigation.");

    Native.SetForegroundWindow((IntPtr)Field(form,"masterHandle"));Pump();

    var showField=typeof(MainForm).GetField("masterShow",BindingFlags.NonPublic|BindingFlags.Instance);showField.SetValue(form,new BusyShow());

    try{Call(form,"CheckSession");if((bool)Field(form,"paused")||!(bool)Field(form,"running"))throw new Exception("Transient COM failure paused the session");}finally{showField.SetValue(form,master);}

    master.View.GotoSlide(2,-1);follower.View.GotoSlide(1,-1);Call(form,"CheckSession");Pump();

    if((int)follower.View.Slide.SlideIndex!=2)throw new Exception("Automatic retry did not recover");

    Console.WriteLine("PASS: injected busy COM call does not pause; next poll restores synchronization.");

   }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}

   finally{Call(form,"Stop");form.Close();}

  }

 }

 static bool CallActive(MainForm form){return (bool)typeof(MainForm).GetMethod("Active",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);}

}

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
class EndOfDeckTest {
 static object Get(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){try{typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}catch(TargetInvocationException ex){throw ex.InnerException;}}
 static void Assert(bool value,string message){if(!value)throw new Exception(message);}
 static void Pump(){for(int i=0;i<10;i++){Application.DoEvents();Thread.Sleep(20);}}
 [STAThread]static void Main(){
  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);
  string folder=Path.GetFullPath(".build/end-deck-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",folder);
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
  try{
   int baseline=(int)app.Presentations.Count;
   foreach(string name in new[]{"presenter","audience"}){dynamic deck=app.Presentations.Add(-1);try{for(int page=1;page<=2;page++){dynamic slide=deck.Slides.Add(page,12);for(int i=0;i<2;i++){dynamic shape=slide.Shapes.AddShape(1,100+i*100,100,80,80);slide.TimeLine.MainSequence.AddEffect(shape,1,0,1);}}deck.SaveAs(Path.Combine(folder,name+".pptx"));}finally{deck.Saved=-1;deck.Close();}}
   using(var f=new MainForm()){
    f.Show();((ComboBox)Get(f,"masterPath")).Text=Path.Combine(folder,"presenter.pptx");((ComboBox)Get(f,"followerPath")).Text=Path.Combine(folder,"audience.pptx");Call(f,"Start");
    dynamic source=Get(f,"masterShow"),audience=Get(f,"followerShow");source.View.Last();source.View.Next();source.View.Next();Call(f,"MatchNativeState");Pump();
    Assert((int)source.View.GetClickIndex()==2&&(int)audience.View.GetClickIndex()==2,"Final animations not synchronized");
    for(int i=0;i<5;i++)Call(f,"AdvancePresenter");
    Assert((int)source.View.State==1&&(int)source.View.Slide.SlideIndex==2,"Repeated forward commands left last slide");
    source.View.Next();Assert((int)source.View.State==5,"Native end screen not reached for recovery test");Call(f,"MatchNativeState");
    Assert((bool)Get(f,"running")&&(int)source.View.State==1&&(int)source.View.GetClickIndex()==2&&(int)audience.View.GetClickIndex()==2,"End screen closed/reset pair");
    source.View.Previous();Call(f,"MatchNativeState");Pump();Assert((int)source.View.GetClickIndex()==(int)audience.View.GetClickIndex(),"Backward navigation after end diverged");
    Call(f,"AdvancePresenter");Call(f,"MatchNativeState");Pump();Assert((int)source.View.GetClickIndex()==2&&(int)audience.View.GetClickIndex()==2,"Forward navigation after end diverged");
    Native.SetForegroundWindow((IntPtr)Get(f,"masterHandle"));Pump();
    IntPtr mouseMemory=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Native.Mouse)));try{
     var bounds=Native.Client((IntPtr)Get(f,"masterHandle"));Marshal.StructureToPtr(new Native.Mouse{Point=new System.Drawing.Point(bounds.X+bounds.Width/2,bounds.Y+bounds.Height/2)},mouseMemory,false);
     var mouse=typeof(MainForm).GetMethod("MouseInput",BindingFlags.NonPublic|BindingFlags.Instance);
     Assert((IntPtr)mouse.Invoke(f,new object[]{0,new IntPtr(0x201),mouseMemory})==new IntPtr(1),"Final-slide click was not suppressed");
     Assert((IntPtr)mouse.Invoke(f,new object[]{0,new IntPtr(0x202),mouseMemory})==new IntPtr(1),"Final-slide click release was not suppressed");
     typeof(MainForm).GetMethod("QueueWheel",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,new object[]{-120});Assert((int)source.View.State==1,"Wheel advanced past final slide");
    }finally{Marshal.FreeHGlobal(mouseMemory);}
    IntPtr memory=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Native.Key)));try{
     Marshal.StructureToPtr(new Native.Key{Code=39},memory,false);var keyboard=typeof(MainForm).GetMethod("Keyboard",BindingFlags.NonPublic|BindingFlags.Instance);var result=(IntPtr)keyboard.Invoke(f,new object[]{0,new IntPtr(0x100),memory});Assert(result==new IntPtr(1),"Forward key escaped guarded navigation");Pump();Assert((bool)Get(f,"running"),"Forward key closed pair");
     Marshal.StructureToPtr(new Native.Key{Code=27},memory,false);result=(IntPtr)keyboard.Invoke(f,new object[]{0,new IntPtr(0x100),memory});Assert(result==new IntPtr(1),"Escape not handled");
    }finally{Marshal.FreeHGlobal(memory);}
    var end=DateTime.UtcNow.AddSeconds(5);while((bool)Get(f,"running")||(bool)Get(f,"stopping")){if(DateTime.UtcNow>end)throw new Exception("Escape did not stop pair");Pump();}
    Assert((int)app.Presentations.Count==baseline,"Escape did not close owned decks");f.Close();
   }
   Console.WriteLine("PASS: final animations retained, repeated forward commands clamped, native end screen recovered, backward/forward navigation retained, Escape closed pair.");
  }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}finally{if((int)app.Presentations.Count==0)app.Quit();}
 }
}

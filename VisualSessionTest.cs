using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
class VisualSessionTest {
 static object Field(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 [STAThread]static void Main(){
  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/visual-settings"));
  using(var f=new MainForm()){
   f.Shown+=(sender,e)=>f.BeginInvoke((Action)(()=>{
    try{
     ((ComboBox)Field(f,"masterPath")).Text=Path.GetFullPath("test-master.pptx");((ComboBox)Field(f,"followerPath")).Text=Path.GetFullPath("test-follower.pptx");
     for(int i=0;i<Screen.AllScreens.Length;i++)if(Screen.AllScreens[i].Primary)((ComboBox)Field(f,"masterMonitor")).SelectedIndex=i;
     ((CheckBox)Field(f,"laser")).Checked=true;Call(f,"Start");
     foreach(dynamic deck in new[]{Field(f,"master"),Field(f,"follower")})foreach(dynamic slide in deck.Slides){float w=(float)deck.PageSetup.SlideWidth,h=(float)deck.PageSetup.SlideHeight;foreach(float x in new[]{0f,w-30})foreach(float y in new[]{0f,h-30}){var q=slide.Shapes.AddShape(1,x,y,30,30);q.Fill.ForeColor.RGB=255;}}
     dynamic a=Field(f,"masterShow"),b=Field(f,"followerShow");a.View.GotoSlide(2,-1);b.View.GotoSlide(2,-1);a.View.GotoSlide(1,-1);b.View.GotoSlide(1,-1);
     Console.WriteLine("VISUAL presenter="+Field(f,"masterHandle")+" audience="+Field(f,"followerHandle"));
    }catch(Exception ex){Console.WriteLine(ex);Call(f,"Stop");f.Close();Environment.ExitCode=1;}
   }));
   var timer=new Timer{Interval=2000};int elapsed=0;timer.Tick+=(s,e)=>{elapsed+=2;if(!(bool)Field(f,"running")&&!(bool)Field(f,"stopping")){timer.Stop();f.Close();return;}var p=(PointerForm)Field(f,"pointer");Console.WriteLine("Pointer visible="+p.Visible+" laser="+p.Laser+" error="+p.Error);if(elapsed>=180){timer.Stop();Call(f,"Stop");f.Close();}};timer.Start();Application.Run(f);timer.Dispose();
  }
 }
}

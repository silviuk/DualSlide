using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
class AppearVisualTest {
 static object Field(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 [STAThread]static void Main(string[] args){
  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));Application.EnableVisualStyles();
  AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/appear-settings"));
  string folder=Path.GetFullPath(".build/appear-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
  dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
  foreach(string side in new[]{"presenter","audience"}){
   dynamic deck=app.Presentations.Add(-1);
   try{deck.PageSetup.SlideWidth=960;deck.PageSetup.SlideHeight=540;
    for(int page=1;page<=2;page++){
     dynamic slide=deck.Slides.Add(page,12);slide.FollowMasterBackground=0;slide.Background.Fill.Solid();slide.Background.Fill.ForeColor.RGB=0xffffff;
     var title=slide.Shapes.AddTextbox(1,45,30,860,70);title.TextFrame.TextRange.Text=side.ToUpper()+" — APPEAR TEST "+page;title.TextFrame.TextRange.Font.Size=28;title.TextFrame.TextRange.Font.Color.RGB=0;
     for(int i=0;i<3;i++){
      dynamic shape=slide.Shapes.AddShape(1,60+i*300,180,240,200);shape.Fill.Solid();shape.Fill.ForeColor.RGB=new[]{0x0000ff,0x008000,0xff0000}[i];shape.Line.Visible=0;
      shape.TextFrame.TextRange.Text="APPEAR "+(i+1);shape.TextFrame.TextRange.Font.Size=28;shape.TextFrame.TextRange.Font.Color.RGB=0xffffff;
      slide.TimeLine.MainSequence.AddEffect(shape,1,0,1);
     }
    }
    deck.SlideShowSettings.ShowWithAnimation=-1;deck.SaveAs(Path.Combine(folder,side+".pptx"));
   }finally{deck.Saved=-1;deck.Close();}
  }
  using(var f=new MainForm()){
   f.Shown+=(s,e)=>f.BeginInvoke((Action)(()=>{
    try{((ComboBox)Field(f,"masterPath")).Text=Path.Combine(folder,"presenter.pptx");((ComboBox)Field(f,"followerPath")).Text=Path.Combine(folder,"audience.pptx");
     for(int i=0;i<Screen.AllScreens.Length;i++)if(Screen.AllScreens[i].Primary)((ComboBox)Field(f,"masterMonitor")).SelectedIndex=i;
     Call(f,"Start");Console.WriteLine("PRESENTER="+Field(f,"masterHandle")+" AUDIENCE="+Field(f,"followerHandle"));
     var advance=new Timer{Interval=3000};advance.Tick+=(a,b)=>{advance.Stop();advance.Dispose();dynamic show=Field(f,"masterShow");int steps=args.Length>0?3:1;for(int i=0;i<steps;i++){show.View.Next();Call(f,"CheckSession");}if(args.Length>0&&args[0]=="back"){show.View.GotoSlide(1,-1);show.View.Next();Call(f,"CheckSession");}Console.WriteLine("Native presenter steps applied; mode="+(args.Length>0?args[0]:"first"));};advance.Start();
    }catch(Exception ex){Console.WriteLine(ex);Call(f,"Stop");f.Close();Environment.ExitCode=1;}
   }));
   var timer=new Timer{Interval=2000};int seconds=0;timer.Tick+=(s,e)=>{seconds+=2;if(!(bool)Field(f,"running")&&!(bool)Field(f,"stopping")){timer.Stop();f.Close();return;}try{dynamic a=Field(f,"masterShow"),b=Field(f,"followerShow");Console.WriteLine("STATE "+a.View.Slide.SlideIndex+"/"+a.View.GetClickIndex()+" "+b.View.Slide.SlideIndex+"/"+b.View.GetClickIndex());}catch{}if(seconds>=150){timer.Stop();Call(f,"Stop");f.Close();}};timer.Start();Application.Run(f);timer.Dispose();
  }
 }
}

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;

class KeepOpenLiveTest {
 static object Field(MainForm f,string n){return typeof(MainForm).GetField(n,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string n){typeof(MainForm).GetMethod(n,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}
 static void Assert(bool b,string text){if(!b)throw new Exception(text);}
 static string Hash(string path){using(var sha=SHA256.Create())using(var file=File.OpenRead(path))return Convert.ToBase64String(sha.ComputeHash(file));}
 static void Pump(){Application.DoEvents();Thread.Sleep(25);}
 [STAThread]static void Main(){
  Native.SetProcessDpiAwarenessContext(new IntPtr(-4));Application.EnableVisualStyles();
  string root=Path.GetFullPath(".build/keep-live-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
  AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",root);AppDomain.CurrentDomain.SetData("DualSlide.ThrowStartupErrors",true);
  string a=Path.Combine(root,"presenter.pptx"),b=Path.Combine(root,"audience.pptx");dynamic app=null;
  try{
   Assert(Screen.AllScreens.Length>=2,"Two extended displays are required");app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
   foreach(string path in new[]{a,b}){dynamic deck=app.Presentations.Add(-1);deck.Slides.Add(1,12);deck.SaveAs(path,24);deck.Close();}
   string ah=Hash(a),bh=Hash(b);DateTime at=File.GetLastWriteTimeUtc(a),bt=File.GetLastWriteTimeUtc(b);
   using(var form=new MainForm()){
    form.Show();((ComboBox)Field(form,"masterPath")).Text=a;((ComboBox)Field(form,"followerPath")).Text=b;((CheckBox)Field(form,"keepOpen")).Checked=true;
    for(int mode=0;mode<3;mode++){
     Call(form,"Start");Assert((bool)Field(form,"running"),"Start failed");
     Assert((int)((dynamic)Field(form,"master")).ReadOnly!=0&&(int)((dynamic)Field(form,"follower")).ReadOnly!=0,"Active shows must be read-only");
     var settle=DateTime.UtcNow.AddSeconds(1);while(DateTime.UtcNow<settle)Pump();
     if(mode==0)Call(form,"Stop");else{dynamic show=Field(form,mode==1?"masterShow":"followerShow");show.View.Exit();Call(form,"CheckSession");}
     var deadline=DateTime.UtcNow.AddSeconds(15);while((bool)Field(form,"stopping")&&DateTime.UtcNow<deadline)Pump();Assert(!(bool)Field(form,"stopping"),"Stop did not finish");
     foreach(string path in new[]{a,b}){
      dynamic reopened=null;foreach(dynamic deck in app.Presentations)if(String.Equals((string)deck.FullName,path,StringComparison.OrdinalIgnoreCase))reopened=deck;
      Assert(reopened!=null,"Presentation was not kept open");Assert((int)reopened.ReadOnly==0,"Reopened presentation is not editable");Assert((int)reopened.Saved==-1,"Reopened original is dirty");
      reopened.Close();
      using(var file=new FileStream(path,FileMode.Open,FileAccess.Write,FileShare.ReadWrite)){}
     }
     Assert(Hash(a)==ah&&Hash(b)==bh&&File.GetLastWriteTimeUtc(a)==at&&File.GetLastWriteTimeUtc(b)==bt,"Original bytes or timestamps changed");
     Console.WriteLine("PASS: "+new[]{"Stop button","presenter End Show","audience End Show"}[mode]+" leaves both original decks editable, clean, unchanged, with write guards released.");
    }
    form.Close();
   }
  }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}
  finally{if(app!=null)foreach(string path in new[]{a,b})try{foreach(dynamic deck in app.Presentations)if(String.Equals((string)deck.FullName,path,StringComparison.OrdinalIgnoreCase)){deck.Saved=-1;deck.Close();break;}}catch{}}
 }
}

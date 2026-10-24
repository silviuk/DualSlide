using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
class FeatureCheck {
 static object Field(MainForm f,string name){return typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
 static void Call(MainForm f,string name,params object[] args){typeof(MainForm).GetMethod(name,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,args);}
 static void Assert(bool b,string message){if(!b)throw new Exception(message);}
 static void Pump(int seconds){var end=DateTime.UtcNow.AddSeconds(seconds);while(DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(10);}}
 [STAThread]static void Main(string[] args){Native.SetProcessDpiAwarenessContext(new IntPtr(-4));Application.EnableVisualStyles();string root=Path.GetFullPath(".build/feature-settings");Directory.CreateDirectory(root);AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",root);
 try{string a=Path.GetFullPath("test-master.pptx"),b=Path.GetFullPath("test-follower.pptx"),c=Path.GetFullPath("chain-master.pptx");var history=new DeckHistory(Path.Combine(root,"DualSlide.history.xml"));history.Remember(a,b);history.Remember(a,c);history=new DeckHistory(Path.Combine(root,"DualSlide.history.xml"));Assert(history.Partner(a)==c&&history.Partner(c)==a,"Bidirectional/latest pair not persisted");Assert(history.Files()[0]==a&&history.Files()[1]==c,"Recent order lost");history.Remember(a,b);Console.WriteLine("PASS: recent pairs persist, latest partner wins, reverse lookup works.");
 using(var form=new MainForm()){form.Show();Pump(1);var presenter=(ComboBox)Field(form,"masterPath");var audience=(ComboBox)Field(form,"followerPath");presenter.Text=a;Call(form,"Associate",presenter);Assert(audience.Text==b,"Presenter selection did not choose paired audience");audience.Text=a;Call(form,"Associate",audience);Assert(presenter.Text==b,"Audience selection did not choose paired presenter");Assert(presenter.Items.Contains(a)&&audience.Items.Contains(b),"Recent dropdowns missing files");Console.WriteLine("PASS: both recent-file dropdowns restore the associated opposite deck.");
 var uiHistory=(DeckHistory)Field(form,"history");uiHistory.Remember(a,c);audience.Text=b;Call(form,"Associate",audience);Call(form,"ValidateAssociation",presenter);Assert(presenter.Text==a&&audience.Text==b,"Focus-only validation changed the chosen pair");uiHistory.Remember(a,b);Console.WriteLine("PASS: moving focus without editing does not replace the chosen pair.");
 var pm=(ComboBox)Field(form,"masterMonitor");var am=(ComboBox)Field(form,"followerMonitor");pm.SelectedIndex=1;Assert(am.SelectedIndex==0,"Presenter collision not resolved");am.SelectedIndex=1;Assert(pm.SelectedIndex==0,"Audience collision not resolved");Console.WriteLine("PASS: display collisions resolve in either direction.");presenter.Text=a;audience.Text=b;
 foreach(bool dark in new[]{false,true}){ModernWindows.Apply(form,dark);Pump(1);foreach(string name in new[]{"start","pause","stop","resync"}){var button=(ModernButton)Field(form,name);using(var bitmap=new Bitmap(button.Width,button.Height)){button.DrawToBitmap(bitmap,button.ClientRectangle);Assert(bitmap.GetPixel(0,0).ToArgb()==button.Parent.BackColor.ToArgb(),"Button corner artifact: "+name);}}Console.WriteLine("PASS: clean button corners, "+(dark?"dark":"light")+"; accent="+ModernWindows.Accent);Pump(args.Length>0?20:0);}
 form.Close();}
 using(var form=new MainForm()){Assert(((ComboBox)Field(form,"masterPath")).Text==a&&((ComboBox)Field(form,"followerPath")).Text==b,"Last selections lost after restart");Assert(((ComboBox)Field(form,"masterMonitor")).SelectedIndex!=((ComboBox)Field(form,"followerMonitor")).SelectedIndex,"Distinct displays lost after restart");Console.WriteLine("PASS: deck and display selections survive restart.");form.Close();}
 }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}}
}

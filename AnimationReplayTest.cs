using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
public sealed class ReplaySlide {public int SlideIndex=1;}
public sealed class ReplaySlides {public int Count=2;}
public sealed class ReplayDeck {public ReplaySlides Slides=new ReplaySlides();}
public sealed class ReplayShow {public ReplayView View=new ReplayView();}
public sealed class ReplayView {
 public ReplaySlide Slide=new ReplaySlide();public int State=1,Index,Rendered,Resets,NextCalls,GotoClickCalls;public bool Busy;
 public int GetClickIndex(){return Index;}public int GetClickCount(){return 12;}
 public void GotoSlide(int slide,int reset){Slide.SlideIndex=slide;Index=Rendered=0;Resets++;}
 // Reproduce Office's problematic jump: reported state moves, rendered state does not.
 public void GotoClick(int click){Index=click;GotoClickCalls++;}
 public void Last(){Slide.SlideIndex=2;State=1;}
 public void Next(){if(Busy){Busy=false;throw new COMException("Office busy",unchecked((int)0x80010001));}NextCalls++;Index++;Rendered=Index;}
}
class AnimationReplayTest {
 static void Set(MainForm f,string name,object value){typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(f,value);}
 static void Match(MainForm f){try{typeof(MainForm).GetMethod("MatchNativeState",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(f,null);}catch(TargetInvocationException ex){throw ex.InnerException;}}
 static void Assert(bool value,string text){if(!value)throw new Exception(text);}
 [STAThread]static void Main(){AppDomain.CurrentDomain.SetData("DualSlide.SettingsDirectory",Path.GetFullPath(".build/replay-settings"));
  using(var form=new MainForm())try{
   var source=new ReplayShow();var audience=new ReplayShow();Set(form,"master",new ReplayDeck());Set(form,"masterShow",source);Set(form,"followerShow",audience);
   source.View.Index=3;Match(form);Assert(audience.View.Rendered==3&&audience.View.GotoClickCalls==0,"Forward Appear content not rendered");
   source.View.Index=1;Match(form);Assert(audience.View.Rendered==1&&audience.View.Resets==1,"Backward replay left wrong visible content");
   source.View.Index=0;Match(form);Assert(audience.View.Rendered==0,"Reset did not hide entrance objects");
   source.View.Slide.SlideIndex=2;source.View.Index=10;Match(form);Assert(audience.View.Rendered==8,"Catch-up must be bounded per poll");Match(form);Assert(audience.View.Rendered==10&&audience.View.Slide.SlideIndex==2,"Multi-step catch-up lost content");
   source.View.Index=11;audience.View.Busy=true;bool rejected=false;try{Match(form);}catch(COMException){rejected=true;}Assert(rejected&&audience.View.Rendered==10,"Busy call changed rendered state");Match(form);Assert(audience.View.Rendered==11,"Busy call did not retry");
   source.View.Index=12;Match(form);int calls=source.View.NextCalls;typeof(MainForm).GetMethod("AdvancePresenter",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);Assert(source.View.NextCalls==calls,"Forward at final build advanced beyond the deck");
   source.View.State=5;Match(form);Assert(source.View.State==1&&audience.View.State==1&&audience.View.Rendered==12,"End screen did not retain final slide/build");
   source.View.Index=11;Match(form);Match(form);Assert(audience.View.Rendered==11,"Cannot go backwards after reaching end");
   Assert(audience.View.GotoClickCalls==0,"Ordinary animations used the non-rendering jump path");Console.WriteLine("PASS: animation replay, busy retry, final-build forward clamp, end-screen recovery, and backward navigation after end.");
  }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}finally{Set(form,"master",null);((PointerForm)typeof(MainForm).GetField("pointer",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form)).Dispose();}
 }
}

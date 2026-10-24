using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

// Office marks HWND restricted: it rejects IDispatch name/Invoke access.
// Use the stable dual-interface vtable, in Office type-library member order.
[ComImport,Guid("91493453-5A91-11CF-8700-00AA0060263B"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
interface ISlideShowHandle {
    object Application {get;} object Parent {get;} object View {get;} object Presentation {get;}
    int IsFullScreen {get;}
    float Left {get;set;} float Top {get;set;} float Width {get;set;} float Height {get;set;}
    int HWND {get;}
}

// One controller per Windows logon session, including executables in other folders.
sealed class SingleInstance : IDisposable {
    readonly Mutex mutex; readonly bool owned;
    public static readonly int ActivateMessage=RegisterWindowMessage("SilviuVlasceanu.DualSlide.Activate.v1");
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int RegisterWindowMessage(string name);
    public SingleInstance(){mutex=new Mutex(false,@"Local\SilviuVlasceanu.DualSlide.Controller");try{owned=mutex.WaitOne(0);}catch(AbandonedMutexException){owned=true;}}
    public bool IsFirst{get{return owned;}}
    public void Notify(){Native.PostMessage(new IntPtr(0xffff),ActivateMessage,IntPtr.Zero,IntPtr.Zero);}
    public void Dispose(){if(owned)mutex.ReleaseMutex();mutex.Dispose();}
}

static class SessionLog {
    static readonly object gate=new object();
    public static void Write(string text){try{lock(gate){Directory.CreateDirectory(Preferences.Root);string path=Path.Combine(Preferences.Root,"DualSlide.log");if(File.Exists(path)&&new FileInfo(path).Length>1024*1024)File.Move(path,path+"."+DateTime.UtcNow.Ticks);File.AppendAllText(path,DateTime.UtcNow.ToString("o")+" "+text+Environment.NewLine);}}catch{}}
}

// Low-level hooks must pump independently of Office COM. Never wait for the UI here.
sealed class InputPump : IDisposable {
    readonly Thread thread; readonly ManualResetEvent ready=new ManualResetEvent(false);
    readonly Native.Hook keyboard,mouse; Control dispatch;
    public IntPtr KeyboardHandle,MouseHandle;
    Exception failure;volatile bool stopping;
    public InputPump(Native.Hook key,Native.Hook pointer){keyboard=key;mouse=pointer;thread=new Thread(Run){IsBackground=true,Name="DualSlide input"};thread.SetApartmentState(ApartmentState.STA);thread.Start();if(!ready.WaitOne(5000)){Dispose();throw new Exception("Input thread did not start.");}if(failure!=null)throw new Exception("Input hooks could not start.",failure);}
    void Run(){try{dispatch=new Control();var handle=dispatch.Handle;if(stopping)return;var module=Native.GetModuleHandle(null);KeyboardHandle=Native.SetWindowsHookEx(13,keyboard,module,0);MouseHandle=Native.SetWindowsHookEx(14,mouse,module,0);if(KeyboardHandle==IntPtr.Zero||MouseHandle==IntPtr.Zero)throw new Exception("Windows rejected an input hook.");ready.Set();if(!stopping)Application.Run();}catch(Exception ex){failure=ex;ready.Set();}finally{Unhook();if(dispatch!=null)dispatch.Dispose();}}
    void Unhook(){var key=Interlocked.Exchange(ref KeyboardHandle,IntPtr.Zero);var pointer=Interlocked.Exchange(ref MouseHandle,IntPtr.Zero);if(key!=IntPtr.Zero)Native.UnhookWindowsHookEx(key);if(pointer!=IntPtr.Zero)Native.UnhookWindowsHookEx(pointer);}
    public void Dispose(){// Unhook immediately; do not wait for an Office or UI thread.
        stopping=true;Unhook();
        if(dispatch!=null&&!dispatch.IsDisposed)try{dispatch.BeginInvoke((Action)Application.ExitThread);}catch(InvalidOperationException){}
    }
}

// Recovery is deliberately independent of both COM and the controller's message pump.
sealed class SessionWatchdog : IDisposable {
    readonly System.Threading.Timer timer;readonly uint uiThread;readonly Action recovered;
    readonly int timeout;long heartbeat;int tripped;volatile bool armed;bool cancellationEnabled;
    IntPtr presenter,audience;readonly IntPtr token=new IntPtr(Guid.NewGuid().GetHashCode()|1);
    const string OwnerProperty="DualSlide.SessionWatchdog.Owner";
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern bool SetProp(IntPtr h,string name,IntPtr value);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern IntPtr GetProp(IntPtr h,string name);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern IntPtr RemoveProp(IntPtr h,string name);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("ole32.dll")] static extern int CoEnableCallCancellation(IntPtr reserved);
    [DllImport("ole32.dll")] static extern int CoDisableCallCancellation(IntPtr reserved);
    [DllImport("ole32.dll")] static extern int CoCancelCall(uint thread,uint timeoutSeconds);
    [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr h,int command);
    public SessionWatchdog(Action recovery):this(recovery,15000){}
    internal SessionWatchdog(Action recovery,int milliseconds){recovered=recovery;timeout=milliseconds;uiThread=GetCurrentThreadId();cancellationEnabled=CoEnableCallCancellation(IntPtr.Zero)>=0;Beat();timer=new System.Threading.Timer(Check,null,250,250);}
    public void Beat(){Interlocked.Exchange(ref heartbeat,DateTime.UtcNow.Ticks);}
    public void Track(IntPtr master,IntPtr follower){if(!cancellationEnabled)cancellationEnabled=CoEnableCallCancellation(IntPtr.Zero)>=0;presenter=master;audience=follower;foreach(var h in new[]{master,follower})if(h!=IntPtr.Zero)SetProp(h,OwnerProperty,token);Beat();armed=true;}
    public void Disarm(){armed=false;foreach(var h in new[]{presenter,audience})if(h!=IntPtr.Zero&&GetProp(h,OwnerProperty)==token)RemoveProp(h,OwnerProperty);Interlocked.Exchange(ref tripped,0);Beat();}
    void Check(object unused){if(!armed||TimeSpan.FromTicks(DateTime.UtcNow.Ticks-Interlocked.Read(ref heartbeat)).TotalMilliseconds<timeout||Interlocked.Exchange(ref tripped,1)!=0)return;
        // Asynchronous window operations never wait on PowerPoint's input queue.
        foreach(var h in new[]{presenter,audience})if(h!=IntPtr.Zero&&Native.IsWindow(h)&&GetProp(h,OwnerProperty)==token){Native.SetWindowPos(h,new IntPtr(-2),0,0,0,0,0x4213);ShowWindowAsync(h,6);}
        SessionLog.Write("PowerPoint/controller stopped responding. Released fullscreen windows and requested COM cancellation.");
        recovered();CoCancelCall(uiThread,0);
    }
    public void Dispose(){Disarm();timer.Dispose();if(cancellationEnabled)CoDisableCallCancellation(IntPtr.Zero);}
}

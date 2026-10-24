using System;

using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Point = System.Drawing.Point;
using Rectangle = System.Drawing.Rectangle;
using MediaColor = System.Windows.Media.Color;

// WPF owns per-pixel alpha composition. The render thread never calls PowerPoint.
sealed class PointerForm : IDisposable {
    sealed class State {public Rectangle Source,Target;public IntPtr Master,Audience;public bool Enabled,Laser;}
    volatile State state=new State();
    readonly Thread thread;readonly ManualResetEvent ready=new ManualResetEvent(false);
    Dispatcher dispatcher;Window window;IntPtr handle;volatile bool visible;volatile string error;
    long ticks,uploads;public bool Visible{get{return visible;}}public bool Laser{get{return state.Laser;}}
    public Rectangle Bounds{get{return state.Target;}}public string Error{get{return error;}}
    public long FrameTicks{get{return Interlocked.Read(ref ticks);}}public long SpriteUploads{get{return Interlocked.Read(ref uploads);}}
    public PointerForm(){thread=new Thread(Run){IsBackground=true,Name="DualSlide pointer"};thread.SetApartmentState(ApartmentState.STA);thread.Start();ready.WaitOne();}
    public void Configure(Rectangle source,Rectangle target,IntPtr master,bool laser){Configure(source,target,master,IntPtr.Zero,laser);}
    public void Configure(Rectangle source,Rectangle target,IntPtr master,IntPtr audience,bool laser){state=new State{Source=source,Target=target,Master=master,Audience=audience,Enabled=true,Laser=laser};}
    public void Hide(){state=new State();}
    public void EndSession(){state=new State();if(dispatcher!=null&&!dispatcher.HasShutdownStarted)try{dispatcher.BeginInvoke(new Action(()=>{try{if(window!=null){window.Hide();}}catch(InvalidOperationException){}finally{visible=false;}}));}catch(InvalidOperationException){}}
    [DllImport("user32.dll")] static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    void Run(){
        try{
            dispatcher=Dispatcher.CurrentDispatcher;
            window=new Window{Title="DualSlide pointer",Width=56,Height=56,WindowStyle=WindowStyle.None,ResizeMode=ResizeMode.NoResize,AllowsTransparency=true,Background=Brushes.Transparent,ShowActivated=false,ShowInTaskbar=false,Topmost=true,Focusable=false,IsHitTestVisible=false};
            var laser=LaserArtwork();var arrow=ArrowArtwork();var canvas=new Grid{Width=56,Height=56};canvas.Children.Add(laser);canvas.Children.Add(arrow);window.Content=canvas;
            handle=new WindowInteropHelper(window).EnsureHandle();Native.SetWindowLong(handle,-20,Native.GetWindowLong(handle,-20)|0x080000A0);
            var timer=new DispatcherTimer(DispatcherPriority.Render){Interval=TimeSpan.FromMilliseconds(16)};
            Point previous=new Point(Int32.MinValue,0);int moved=Environment.TickCount;bool lastLaser=false;bool artworkSet=false;Point lastPosition=new Point(Int32.MinValue,0);int lastExtent=0;
            timer.Tick+=(s,e)=>{
                try{
                    Interlocked.Increment(ref ticks);var current=state;Point mouse;GetCursorPos(out mouse);
                    bool isLaser=current.Laser||((Native.GetAsyncKeyState(17)&0x8000)!=0&&(Native.GetAsyncKeyState(1)&0x8000)!=0);
                    if(mouse!=previous){previous=mouse;moved=Environment.TickCount;}
                    if(!current.Enabled||Native.GetForegroundWindow()!=current.Master||current.Source.Width<=0||!current.Source.Contains(mouse)||(!isLaser&&unchecked(Environment.TickCount-moved)>3000)){
                        if(visible){window.Hide();visible=false;}return;
                    }
                    // No foreign HWND owner: owner destruction can destroy our WPF
                    // surface, and ownership couples input queues across processes.
                    if(!artworkSet||lastLaser!=isLaser){laser.Visibility=isLaser?Visibility.Visible:Visibility.Hidden;arrow.Visibility=isLaser?Visibility.Hidden:Visibility.Visible;lastLaser=isLaser;artworkSet=true;Interlocked.Increment(ref uploads);}
                    var position=new Point(current.Target.X+(int)((mouse.X-current.Source.X)*(double)current.Target.Width/current.Source.Width),current.Target.Y+(int)((mouse.Y-current.Source.Y)*(double)current.Target.Height/current.Source.Height));
                    // Show before moving so WPF has actually supplied the translucent surface to DWM.
                    bool justShown=!visible;if(justShown){window.Show();window.UpdateLayout();visible=true;}
                    double scale=GetDpiForWindow(handle)/96.0;int extent=(int)Math.Round(56*scale),half=extent/2;
                    if(justShown||position!=lastPosition||extent!=lastExtent){
                        // This independent click-through surface has no foreign owner.
                        Native.SetWindowPos(handle,new IntPtr(-1),position.X-half,position.Y-half,extent,extent,0x4210);
                        lastPosition=position;lastExtent=extent;
                    }
                    error=null;
                }catch(Exception ex){error=ex.Message;if(visible){window.Hide();visible=false;}}
            };
            ready.Set();timer.Start();Dispatcher.Run();timer.Stop();window.Close();
        }catch(Exception ex){error=ex.Message;ready.Set();}
    }
    static Grid LaserArtwork(){var grid=new Grid{Width=56,Height=56,IsHitTestVisible=false};
        var glow=new RadialGradientBrush();glow.GradientStops.Add(new GradientStop(MediaColor.FromArgb(235,255,12,0),0));glow.GradientStops.Add(new GradientStop(MediaColor.FromArgb(155,255,20,0),0.45));glow.GradientStops.Add(new GradientStop(MediaColor.FromArgb(0,255,20,0),1));glow.Freeze();
        grid.Children.Add(new Ellipse{Width=24,Height=24,Fill=glow});grid.Children.Add(new Ellipse{Width=10,Height=10,Fill=new SolidColorBrush(MediaColor.FromRgb(255,25,8))});grid.Children.Add(new Ellipse{Width=4,Height=4,Fill=new SolidColorBrush(MediaColor.FromArgb(190,255,222,190)),Margin=new Thickness(-2,-3,0,0)});return grid;
    }
    static Grid ArrowArtwork(){var grid=new Grid{Width=56,Height=56,IsHitTestVisible=false};grid.Children.Add(new System.Windows.Shapes.Path{Data=Geometry.Parse("M28,28 L29,48 34,42 40,50 44,48 38,40 47,39 Z"),Fill=Brushes.White,Stroke=Brushes.Black,StrokeThickness=1});return grid;}
    public void Dispose(){if(dispatcher!=null&&!dispatcher.HasShutdownStarted)dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);thread.Join(2000);ready.Dispose();}
}

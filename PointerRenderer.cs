using System;

using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using Point = System.Drawing.Point;
using Rectangle = System.Drawing.Rectangle;

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
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    [StructLayout(LayoutKind.Sequential)] internal struct CursorInfo {public int Size,Flags;public IntPtr Cursor;public Point Position;}
    [StructLayout(LayoutKind.Sequential)] struct IconInfo {public int IsIcon;public uint XHotspot,YHotspot;public IntPtr Mask,Color;}
    [DllImport("user32.dll")] static extern bool GetCursorInfo(ref CursorInfo info);
    [DllImport("user32.dll")] static extern IntPtr CopyIcon(IntPtr icon);
    [DllImport("user32.dll")] static extern bool GetIconInfo(IntPtr icon,out IconInfo info);
    [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr icon);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr value);
    [DllImport("user32.dll")] static extern bool DrawIconEx(IntPtr dc,int x,int y,IntPtr icon,int width,int height,uint step,IntPtr brush,uint flags);
    static byte[] DrawCursor(IntPtr cursor,int width,int height,System.Drawing.Color background){
        using(var bitmap=new System.Drawing.Bitmap(width,height,System.Drawing.Imaging.PixelFormat.Format32bppRgb)){
            using(var graphics=System.Drawing.Graphics.FromImage(bitmap)){graphics.Clear(background);IntPtr dc=graphics.GetHdc();try{if(!DrawIconEx(dc,0,0,cursor,width,height,0,IntPtr.Zero,3))throw new System.ComponentModel.Win32Exception();}finally{graphics.ReleaseHdc(dc);}}
            var data=bitmap.LockBits(new Rectangle(0,0,width,height),System.Drawing.Imaging.ImageLockMode.ReadOnly,System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try{var pixels=new byte[width*height*4];for(int y=0;y<height;y++)Marshal.Copy(IntPtr.Add(data.Scan0,y*data.Stride),pixels,y*width*4,width*4);return pixels;}finally{bitmap.UnlockBits(data);}
        }
    }
    // Copy the borrowed HCURSOR before reading it. GetIconInfo allocates two GDI
    // bitmaps which must be deleted, even when WPF conversion fails.
    internal static BitmapSource CursorArtwork(IntPtr cursor,out Point hotspot){
        hotspot=new Point();IntPtr copy=CopyIcon(cursor);if(copy==IntPtr.Zero)throw new System.ComponentModel.Win32Exception();
        IconInfo info=new IconInfo();
        try{if(!GetIconInfo(copy,out info))throw new System.ComponentModel.Win32Exception();hotspot=new Point((int)info.XHotspot,(int)info.YHotspot);
            var size=Imaging.CreateBitmapSourceFromHIcon(copy,Int32Rect.Empty,BitmapSizeOptions.FromEmptyOptions());
            // WPF's HICON conversion can lose legacy cursor AND-mask alpha.
            // Native composition against black/white recovers both that mask
            // and modern antialiased alpha, without black boxes or jagged edges.
            var dark=DrawCursor(copy,size.PixelWidth,size.PixelHeight,System.Drawing.Color.Black);var light=DrawCursor(copy,size.PixelWidth,size.PixelHeight,System.Drawing.Color.White);
            for(int i=0;i<dark.Length;i+=4){int difference=Math.Max(light[i]-dark[i],Math.Max(light[i+1]-dark[i+1],light[i+2]-dark[i+2]));byte alpha=(byte)(255-Math.Max(0,Math.Min(255,difference)));dark[i+3]=alpha;for(int channel=0;channel<3;channel++)dark[i+channel]=Math.Min(dark[i+channel],alpha);}
            var bitmap=BitmapSource.Create(size.PixelWidth,size.PixelHeight,96,96,PixelFormats.Pbgra32,null,dark,size.PixelWidth*4);bitmap.Freeze();return bitmap;
        }finally{if(info.Mask!=IntPtr.Zero)DeleteObject(info.Mask);if(info.Color!=IntPtr.Zero)DeleteObject(info.Color);DestroyIcon(copy);}
    }
    void Run(){
        try{
            dispatcher=Dispatcher.CurrentDispatcher;
            window=new Window{Title="DualSlide pointer",Width=256,Height=256,WindowStyle=WindowStyle.None,ResizeMode=ResizeMode.NoResize,AllowsTransparency=true,Background=Brushes.Transparent,ShowActivated=false,ShowInTaskbar=false,Topmost=true,Focusable=false,IsHitTestVisible=false};
            var sprite=new Image{Stretch=Stretch.Fill,IsHitTestVisible=false};var canvas=new Canvas{Width=256,Height=256};canvas.Children.Add(sprite);window.Content=canvas;
            var laser=LaserArtwork(255);Canvas.SetLeft(laser,112);Canvas.SetTop(laser,112);canvas.Children.Add(laser);
            handle=new WindowInteropHelper(window).EnsureHandle();Native.SetWindowLong(handle,-20,Native.GetWindowLong(handle,-20)|0x080000A0);
            var timer=new DispatcherTimer(DispatcherPriority.Render){Interval=TimeSpan.FromMilliseconds(16)};
            IntPtr lastCursor=IntPtr.Zero;double lastSourceScale=0;Point lastPosition=new Point(Int32.MinValue,0);int lastExtent=0;
            timer.Tick+=(s,e)=>{
                try{
                    Interlocked.Increment(ref ticks);var current=state;var cursor=new CursorInfo{Size=Marshal.SizeOf(typeof(CursorInfo))};
                    bool available=GetCursorInfo(ref cursor);Point mouse=cursor.Position;
                    bool isLaser=current.Laser||((Native.GetAsyncKeyState(17)&0x8000)!=0&&(Native.GetAsyncKeyState(1)&0x8000)!=0);
                    if(!current.Enabled||!available||(!isLaser&&(cursor.Flags!=1||cursor.Cursor==IntPtr.Zero))||Native.GetForegroundWindow()!=current.Master||current.Source.Width<=0||!current.Source.Contains(mouse)){
                        if(visible){window.Hide();visible=false;}return;
                    }
                    // No foreign HWND owner: owner destruction can destroy our WPF
                    // surface, and ownership couples input queues across processes.
                    double sourceScale=Math.Max(96,GetDpiForWindow(current.Master))/96.0;
                    laser.Visibility=isLaser?Visibility.Visible:Visibility.Hidden;sprite.Visibility=isLaser?Visibility.Hidden:Visibility.Visible;
                    if(!isLaser&&(lastCursor!=cursor.Cursor||lastSourceScale!=sourceScale)){Point hotspot;var bitmap=CursorArtwork(cursor.Cursor,out hotspot);sprite.Source=bitmap;
                        sprite.Width=bitmap.PixelWidth/sourceScale;sprite.Height=bitmap.PixelHeight/sourceScale;
                        Canvas.SetLeft(sprite,128-hotspot.X/sourceScale);Canvas.SetTop(sprite,128-hotspot.Y/sourceScale);
                        lastCursor=cursor.Cursor;lastSourceScale=sourceScale;Interlocked.Increment(ref uploads);}
                    var position=new Point(current.Target.X+(int)((mouse.X-current.Source.X)*(double)current.Target.Width/current.Source.Width),current.Target.Y+(int)((mouse.Y-current.Source.Y)*(double)current.Target.Height/current.Source.Height));
                    // Show before moving so WPF has actually supplied the translucent surface to DWM.
                    bool justShown=!visible;if(justShown){window.Show();window.UpdateLayout();visible=true;}
                    double scale=Math.Max(96,GetDpiForWindow(handle))/96.0;int extent=(int)Math.Round(256*scale),half=extent/2;
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
    internal static Grid LaserArtwork(int rgb){
        // PowerPoint renders its laser in the slide surface (HCURSOR is null).
        // A small saturated ring and translucent center sit inside a soft halo.
        var color=Color.FromRgb((byte)rgb,(byte)(rgb>>8),(byte)(rgb>>16));
        var halo=new RadialGradientBrush();halo.GradientStops.Add(new GradientStop(Color.FromArgb(160,color.R,color.G,color.B),0));halo.GradientStops.Add(new GradientStop(Color.FromArgb(75,color.R,color.G,color.B),0.35));halo.GradientStops.Add(new GradientStop(Color.FromArgb(0,color.R,color.G,color.B),1));halo.Freeze();
        var grid=new Grid{Width=32,Height=32,IsHitTestVisible=false};
        grid.Children.Add(new Ellipse{Width=32,Height=32,Fill=halo});
        grid.Children.Add(new Ellipse{Width=9,Height=9,Stroke=new SolidColorBrush(color),StrokeThickness=2,Fill=new SolidColorBrush(Color.FromArgb(70,color.R,color.G,color.B))});return grid;
    }
    public void Dispose(){if(dispatcher!=null&&!dispatcher.HasShutdownStarted)dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);thread.Join(2000);ready.Dispose();}
}

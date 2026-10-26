using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
[assembly:System.Reflection.AssemblyVersion("0.5.5.0")]
[assembly:System.Reflection.AssemblyFileVersion("0.5.5.0")]
[assembly:System.Reflection.AssemblyCopyright("(C) 2026 Silviu Vlasceanu")]

static class Program {
    [STAThread] static void Main(string[] args) {
        Native.SetProcessDpiAwarenessContext(new IntPtr(-4));
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        using(var instance=new SingleInstance()){
            if(!instance.IsFirst){instance.Notify();return;}
            Application.ThreadException+=(sender,e)=>SessionLog.Write(e.Exception.ToString());
            AppDomain.CurrentDomain.UnhandledException+=(sender,e)=>SessionLog.Write(Convert.ToString(e.ExceptionObject));
            Application.Run(new MainForm());
        }
    }
}

static class Native {
    public delegate bool EnumWindow(IntPtr hwnd,IntPtr param);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindow callback,IntPtr param);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder text,int count);
    public static List<IntPtr> Shows(){var list=new List<IntPtr>();EnumWindows((h,p)=>{var name=new StringBuilder(256);GetClassName(h,name,256);if(name.ToString().Equals("screenClass",StringComparison.OrdinalIgnoreCase))list.Add(h);return true;},IntPtr.Zero);return list;}
    public static IntPtr NewShow(List<IntPtr> before){foreach(var h in Shows())if(!before.Contains(h))return h;throw new Exception("PowerPoint slideshow window was not found.");}
    public delegate IntPtr Hook(int code, IntPtr message, IntPtr data);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd,int command);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern IntPtr SetWindowsHookEx(int id, Hook callback, IntPtr module, uint thread);
    [DllImport("user32.dll")] public static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hook,int code,IntPtr message,IntPtr data);
    [DllImport("kernel32.dll",CharSet=CharSet.Auto)] public static extern IntPtr GetModuleHandle(string name);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,int message,IntPtr wparam,IntPtr lparam);
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hwnd,out Rect rect);
    [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr hwnd,ref Point point);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hwnd);
    [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd,int name,ref int value,int size);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hwnd,IntPtr after,int x,int y,int w,int h,uint flags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hwnd,int index);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hwnd,int index,int value);
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int key);
    public struct Rect { public int Left,Top,Right,Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Mouse { public Point Point; public uint Data,Flags,Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)] public struct Key { public uint Code,Scan,Flags,Time; public UIntPtr Extra; }
    public static Rectangle Client(IntPtr h) {
        Rect r; GetClientRect(h,out r); Point p=new Point(); ClientToScreen(h,ref p);
        return new Rectangle(p.X,p.Y,r.Right-r.Left,r.Bottom-r.Top);
    }
    public static Rectangle Content(Rectangle r,double aspect) {
        int w=r.Width,h=(int)(w/aspect);
        if(h>r.Height){h=r.Height;w=(int)(h*aspect);}
        return new Rectangle(r.X+(r.Width-w)/2,r.Y+(r.Height-h)/2,w,h);
    }
}

sealed class MainForm : Form {
    ComboBox masterPath=new ModernComboBox(), followerPath=new ModernComboBox();
    readonly List<Control> configuration=new List<Control>();
    readonly DeckHistory history=new DeckHistory(System.IO.Path.Combine(Preferences.Root,"DualSlide.history.xml"));
    bool choosingPair,choosingDisplay;
    readonly Queue<int> navigation=new Queue<int>();int wheelRemainder,lastWheel;bool navigating;
    ComboBox masterMonitor=new ModernComboBox(),followerMonitor=new ModernComboBox();
    Button start=new ModernButton(), pause=new ModernButton(),stop=new ModernButton(),resync=new ModernButton();
    Label status=new Label(); CheckBox laser=new CheckBox(),keepOpen=new CheckBox();
    dynamic ppt,master,follower,masterShow,followerShow;
    bool ownPowerPoint;
    bool keepDecksForSession,ownMaster,ownFollower,synchronizing,stopping,cleaning,closeAfterStop,masterConfigured,followerConfigured;
    readonly Queue<string> reopenForEditing=new Queue<string>();string reopenError;
    readonly List<System.IO.FileStream> sourceGuards=new List<System.IO.FileStream>();
    int masterType, followerType, masterPresenter, followerPresenter,masterScrollbar,followerScrollbar;
    IntPtr masterHandle,followerHandle,masterCanvas,followerCanvas,keyHook,mouseHook;
    Native.Hook keyCallback,mouseCallback;
    InputPump input; SessionWatchdog watchdog;bool checking;int recoveryRequested;
    void RecoverDesktop(){running=false;pointer.EndSession();if(input!=null)input.Dispose();System.Threading.Interlocked.Exchange(ref recoveryRequested,1);try{BeginInvoke((Action)(()=>{if(!IsDisposed)Stop();}));}catch(InvalidOperationException){}}

    PointerForm pointer=new PointerForm(); Timer timer=new Timer(),pointerTimer=new Timer();
    bool nativeLaser; Rectangle sourceRect,targetRect;


    volatile bool running,paused; double aspect=16.0/9; uint mouseButtons;
    volatile bool endReached;bool numberEntry,blockedEndClick;
    HashSet<int> held=new HashSet<int>();
    readonly string settingsFile=System.IO.Path.Combine(Preferences.Root,"DualSlide.settings");
    void MatchNativeState(){
        if(synchronizing)return;
        synchronizing=true;
        try{
        int state=(int)masterShow.View.State;
        // The native end screen is not an explicit request to exit. Return to
        // the last slide without restarting it, retaining its animation state.
        if(state==5){masterShow.View.Last();masterShow.View.State=1;state=1;}
        int slide=(int)masterShow.View.Slide.SlideIndex;
        int click=(int)masterShow.View.GetClickIndex();
        // PowerPoint can change slides between COM reads. Never apply a mixed snapshot.
        if((int)masterShow.View.Slide.SlideIndex!=slide)return;
        endReached=PresenterAtEnd();
        if((int)followerShow.View.Slide.SlideIndex!=slide)followerShow.View.GotoSlide(slide,-1);
        if((int)followerShow.View.Slide.SlideIndex!=slide)return;
        if(click>(int)followerShow.View.GetClickCount())return;
        int audienceClick=(int)followerShow.View.GetClickIndex();
        if(audienceClick!=click){
            if(click<0){followerShow.View.GotoClick(click);}
            else{
                // GotoClick can report the requested index without rendering the
                // entrance objects in an inactive windowed show. Play real steps.
                if(audienceClick<0||audienceClick>click){followerShow.View.GotoSlide(slide,-1);audienceClick=(int)followerShow.View.GetClickIndex();}
                for(int budget=0;budget<8&&audienceClick<click;budget++){
                    followerShow.View.Next();
                    if((int)followerShow.View.Slide.SlideIndex!=slide)return;
                    int advanced=(int)followerShow.View.GetClickIndex();
                    if(advanced<=audienceClick)break; // retry on a later poll if Office is still transitioning
                    audienceClick=advanced;
                }
            }
        }
        if((int)followerShow.View.State!=state)followerShow.View.State=state;
        }finally{synchronizing=false;}
    }
    void Command(int key){
        BeginInvoke((Action)(()=>{
            if(!running||paused)return;
            try{
                switch(key){
                    case 13:case 32:case 34:case 39:case 40:
                    case 8:case 33:case 37:case 38:MatchNativeState();break;
                    case 36:case 35:case 66:case 87:MatchNativeState();break;
                    case 27:Stop();break;
                }
            }catch(COMException){status.Text="PowerPoint is busy; synchronization will retry automatically.";}
            catch(Exception ex){status.Text="Waiting for PowerPoint: "+ex.Message;}
        }));
    }
    public MainForm() {
        using(var iconStream=typeof(MainForm).Assembly.GetManifestResourceStream("DualSlide.Icon.ico")){if(iconStream!=null){using(var sourceIcon=new Icon(iconStream)){Icon=(Icon)sourceIcon.Clone();}}}
        Text="DualSlide 0.5.5";Font=new Font("Segoe UI Variable Text",10.5f);AutoScaleMode=AutoScaleMode.Dpi;ClientSize=new Size(850,708);FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;
        var title=new Label{Text="Present together",Font=new Font("Segoe UI Variable Display",25,FontStyle.Bold),AutoSize=true,Location=new Point(28,25)};
        var subtitle=new Label{Text="Two PowerPoint decks, one presentation.",Tag="muted",AutoSize=true,Location=new Point(32,78)};
        Controls.AddRange(new Control[]{title,subtitle});
        AddDeckCard("Presenter",masterPath,masterMonitor,124);
        AddDeckCard("Audience",followerPath,followerMonitor,304);
        foreach(var display in Displays.Read()){masterMonitor.Items.Add(display);followerMonitor.Items.Add(display);}
        masterMonitor.SelectedIndex=0;followerMonitor.SelectedIndex=Screen.AllScreens.Length>1?1:0;
        try{string choiceSource=System.IO.File.Exists(settingsFile)?settingsFile:System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"DualSlide.settings");if(System.IO.File.Exists(choiceSource)){var saved=System.IO.File.ReadAllLines(choiceSource);if(saved.Length>=4){masterPath.Text=saved[0];followerPath.Text=saved[1];SelectDisplay(masterMonitor,saved[2]);SelectDisplay(followerMonitor,saved[3]);keepOpen.Checked=saved.Length>4&&saved[4]==Boolean.TrueString;}}}catch{}
        if(history.Files().Length==0&&System.IO.File.Exists(masterPath.Text)&&System.IO.File.Exists(followerPath.Text))try{history.Remember(masterPath.Text,followerPath.Text);}catch{}
        RefreshHistory();
        if(masterMonitor.SelectedIndex==followerMonitor.SelectedIndex&&masterMonitor.Items.Count>1)followerMonitor.SelectedIndex=masterMonitor.SelectedIndex==0?1:0;
        masterMonitor.SelectedIndexChanged+=(s,e)=>KeepDisplaysDifferent(masterMonitor,followerMonitor);
        followerMonitor.SelectedIndexChanged+=(s,e)=>KeepDisplaysDifferent(followerMonitor,masterMonitor);
        start.Text="Start presentation";pause.Text="Pause";stop.Text="Stop";resync.Text="Reset current slide";
        start.SetBounds(32,550,180,42);pause.SetBounds(224,550,112,42);stop.SetBounds(348,550,112,42);resync.SetBounds(472,550,190,42);
        laser.Text="Start with laser pointer";laser.SetBounds(32,485,520,27);
        keepOpen.Text="Keep presentations open for editing after slideshow";keepOpen.SetBounds(32,515,560,27);keepOpen.AccessibleName=keepOpen.Text;
        keepOpen.CheckedChanged+=(s,e)=>SaveChoices();configuration.Add(keepOpen);
        var identify=new ModernButton{Text="Identify displays"};identify.SetBounds(638,486,180,38);identify.Click+=(s,e)=>Displays.Identify();
        status.SetBounds(32,613,786,40);status.Tag="muted";status.Text="Choose matching decks and two different displays.";
        var copyright=new Label{Text="\u00A9 2026 Silviu Vlasceanu",Tag="muted",AutoSize=true,Location=new Point(32,672),Font=new Font(Font.FontFamily,9)};
        var github=new LinkLabel{Text="github.com/silviuk/DualSlide",AutoSize=true,Location=new Point(586,672),Font=new Font(Font.FontFamily,9),AccessibleName="DualSlide on GitHub"};
        github.LinkClicked+=(s,e)=>{try{Process.Start(new ProcessStartInfo("https://github.com/silviuk/DualSlide"){UseShellExecute=true});}catch(Exception ex){status.Text="Could not open GitHub: "+ex.Message;}};
        Controls.AddRange(new Control[]{start,pause,stop,resync,laser,keepOpen,identify,status,copyright,github});pause.Enabled=stop.Enabled=resync.Enabled=false;
        Shown+=(s,e)=>ModernWindows.Apply(this);
        start.Click+=(s,e)=>Start();stop.Click+=(s,e)=>Stop();
        pause.Click+=(s,e)=>{if(!running)return;paused=!paused;ReleaseInputs();pointer.Hide();pause.Text=paused?"Resume":"Pause";status.Text=paused?"Synchronization paused.":"Synchronized; return to presenter slideshow.";};
        resync.Click+=(s,e)=>{if(!running)return;try{int slide=(int)masterShow.View.Slide.SlideIndex;masterShow.View.GotoSlide(slide,-1);followerShow.View.GotoSlide(slide,-1);status.Text="Both reset to start of slide "+slide;Native.SetForegroundWindow(masterHandle);}catch(Exception ex){Fail(ex.Message);}};
        keyCallback=Keyboard;mouseCallback=MouseInput;watchdog=new SessionWatchdog(RecoverDesktop);
        timer.Interval=50;timer.Tick+=(s,e)=>CheckSession();timer.Start();
        pointerTimer.Interval=16;pointerTimer.Tick+=(s,e)=>UpdatePointer();pointerTimer.Start();
        FormClosing+=(s,e)=>{SaveChoices();Stop();if(stopping){e.Cancel=true;closeAfterStop=true;return;}timer.Dispose();pointerTimer.Dispose();pointer.Dispose();watchdog.Dispose();};
    }
    void AddDeckCard(string title,ComboBox files,ComboBox monitor,int y){
        var card=new ModernCard();card.SetBounds(32,y,786,166);Controls.Add(card);configuration.Add(card);
        var heading=new Label{Text=title,Font=new Font(Font.FontFamily,13,FontStyle.Bold),AutoSize=true,Location=new Point(20,15)};
        files.DropDownStyle=ComboBoxStyle.DropDown;files.FlatStyle=FlatStyle.Flat;files.AutoCompleteMode=AutoCompleteMode.SuggestAppend;files.AutoCompleteSource=AutoCompleteSource.ListItems;files.MaxDropDownItems=10;files.DropDownWidth=1000;files.SetBounds(20,51,612,34);files.AccessibleName=title+" deck";
        files.SelectionChangeCommitted+=(s,e)=>Associate(files);files.Validated+=(s,e)=>ValidateAssociation(files);
        var browse=new ModernButton{Text="Browse..."};browse.SetBounds(644,49,122,36);browse.Click+=(s,e)=>{using(var dialog=new OpenFileDialog{Filter="PowerPoint|*.pptx;*.pptm;*.ppt",CheckFileExists=true})if(dialog.ShowDialog(this)==DialogResult.OK){files.Text=dialog.FileName;Associate(files);}};
        var label=new Label{Text="Display",Tag="muted",AutoSize=true,Location=new Point(20,111)};
        monitor.DropDownStyle=ComboBoxStyle.DropDownList;monitor.FlatStyle=FlatStyle.Flat;monitor.SetBounds(152,105,614,34);monitor.AccessibleName=title+" display";
        card.Controls.AddRange(new Control[]{heading,files,browse,label,monitor});
    }
    void RefreshHistory(){string a=masterPath.Text,b=followerPath.Text;var files=history.Files();foreach(var box in new[]{masterPath,followerPath}){box.BeginUpdate();box.Items.Clear();box.Items.AddRange(files);box.EndUpdate();}masterPath.Text=a;followerPath.Text=b;masterPath.SelectionLength=followerPath.SelectionLength=0;masterPath.Tag=a;followerPath.Tag=b;}
    void ValidateAssociation(ComboBox box){if(!String.Equals(box.Text,box.Tag as string,StringComparison.OrdinalIgnoreCase))Associate(box);}
    void Associate(ComboBox box){if(choosingPair||running)return;box.Tag=box.Text;string path=box.Text.Trim();if(String.IsNullOrEmpty(path))return;try{path=System.IO.Path.GetFullPath(path);}catch{return;}string other=history.Partner(path);if(other==null)return;choosingPair=true;try{box.Text=path;box.Tag=path;var counterpart=box==masterPath?followerPath:masterPath;counterpart.Text=other;counterpart.Tag=other;}finally{choosingPair=false;}}
    void KeepDisplaysDifferent(ComboBox changed,ComboBox other){if(choosingDisplay)return;choosingDisplay=true;try{if(changed.SelectedIndex==other.SelectedIndex&&other.Items.Count>1)for(int i=0;i<other.Items.Count;i++)if(i!=changed.SelectedIndex){other.SelectedIndex=i;break;}}finally{choosingDisplay=false;}SaveChoices();}
    void SaveChoices(){try{if(masterMonitor.SelectedItem!=null&&followerMonitor.SelectedItem!=null)System.IO.File.WriteAllLines(settingsFile,new[]{masterPath.Text,followerPath.Text,((DisplayChoice)masterMonitor.SelectedItem).Device,((DisplayChoice)followerMonitor.SelectedItem).Device,keepOpen.Checked.ToString()});}catch{}}
    static void SelectDisplay(ComboBox box,string id){for(int i=0;i<box.Items.Count;i++){var display=(DisplayChoice)box.Items[i];if(display.Device==id||display.ToString()==id){box.SelectedIndex=i;return;}}}
    protected override void WndProc(ref Message m){
        if(m.Msg==SingleInstance.ActivateMessage){if(running&&Native.IsWindow(masterHandle))Native.SetForegroundWindow(masterHandle);else{if(WindowState==FormWindowState.Minimized)WindowState=FormWindowState.Normal;Show();Activate();}return;}
        base.WndProc(ref m);if((m.Msg==0x1A||m.Msg==0x320||m.Msg==0x31E||m.Msg==0x15)&&IsHandleCreated&&!IsDisposed)BeginInvoke((Action)(()=>{if(!IsDisposed)ModernWindows.Apply(this);}));}
    void Start(){
        if(running||stopping)return;
        reopenError=null;keepDecksForSession=false; // Failed starts always close app-owned decks.
        try {
            SessionLog.Write("Starting presentation session.");
            if(!System.IO.File.Exists(masterPath.Text)||!System.IO.File.Exists(followerPath.Text))throw new Exception("Choose two existing PowerPoint files.");
            if(masterMonitor.SelectedIndex==followerMonitor.SelectedIndex)throw new Exception("Select two different displays.");
            if(String.Equals(masterPath.Text,followerPath.Text,StringComparison.OrdinalIgnoreCase))throw new Exception("Choose two different deck files.");
            masterPath.Text=System.IO.Path.GetFullPath(masterPath.Text);followerPath.Text=System.IO.Path.GetFullPath(followerPath.Text);
            if(String.Equals(masterPath.Text,followerPath.Text,StringComparison.OrdinalIgnoreCase))throw new Exception("Choose two different deck files.");
            ProtectSource(masterPath.Text);ProtectSource(followerPath.Text);
            // PowerPoint can reuse an existing process even through CreateInstance.
            // Be conservative: never Quit if any PowerPoint was already running.
            ownPowerPoint=!PowerPointRunning();
            ppt=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application",true));
            master=OpenDeck(masterPath.Text,out ownMaster);
            follower=OpenDeck(followerPath.Text,out ownFollower);
            masterType=(int)master.SlideShowSettings.ShowType;followerType=(int)follower.SlideShowSettings.ShowType;
            masterPresenter=(int)master.SlideShowSettings.ShowPresenterView;followerPresenter=(int)follower.SlideShowSettings.ShowPresenterView;masterScrollbar=(int)master.SlideShowSettings.ShowScrollbar;followerScrollbar=(int)follower.SlideShowSettings.ShowScrollbar;
            if(master.Slides.Count!=follower.Slides.Count)throw new Exception("Slide counts must match.");
            aspect=(double)master.PageSetup.SlideWidth/(double)master.PageSetup.SlideHeight;
            double other=(double)follower.PageSetup.SlideWidth/(double)follower.PageSetup.SlideHeight;
            if(Math.Abs(aspect-other)>0.001)throw new Exception("Slide aspect ratios must match.");
            masterConfigured=true;masterShow=RunShow(master,SelectedScreen(masterMonitor),out masterHandle,out masterCanvas);
            followerConfigured=true;followerShow=RunShow(follower,SelectedScreen(followerMonitor),out followerHandle,out followerCanvas);

            // Finish audience initialization, then restore presenter focus.
            followerShow.Activate();
            Native.SetWindowPos(followerHandle,IntPtr.Zero,0,0,0,0,0x4213);
            followerShow.View.PointerType=3;
            masterShow.View.LaserPointerEnabled=laser.Checked;
            // These flags describe disposable in-memory show settings, not a disk save.
            if(ownMaster)master.Saved=-1;if(ownFollower)follower.Saved=-1;
            input=new InputPump(keyCallback,mouseCallback);keyHook=input.KeyboardHandle;mouseHook=input.MouseHandle;
            watchdog.Track(masterHandle,followerHandle);
            keepDecksForSession=keepOpen.Checked;
            running=true;paused=false;start.Enabled=false;pause.Enabled=stop.Enabled=resync.Enabled=true;foreach(var control in configuration)control.Enabled=false;status.Text="Running. Input on presenter show is mirrored to audience.";
            try{history.Remember(masterPath.Text,followerPath.Text);RefreshHistory();}catch(Exception ex){status.Text="Running. Recent files could not be saved: "+ex.Message;}
            SaveChoices();
            Native.SetForegroundWindow(masterHandle);
        }catch(Exception ex){SessionLog.Write(ex.ToString());Stop();if(AppDomain.CurrentDomain.GetData("DualSlide.ThrowStartupErrors")!=null)throw;MessageBox.Show(this,ex.Message,"Could not start DualSlide");}
    }
    void ProtectSource(string path){
        try{sourceGuards.Add(new System.IO.FileStream(path,System.IO.FileMode.Open,System.IO.FileAccess.Read,System.IO.FileShare.Read));}
        catch(System.IO.IOException){throw new Exception("Cannot lock the original against changes. Close any editor using this file, then try again: "+path);}
    }
    static bool PowerPointRunning(){var processes=Process.GetProcessesByName("POWERPNT");try{return processes.Length!=0;}finally{foreach(var process in processes)process.Dispose();}}
    static void ReleaseCom(object value){if(value!=null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
    dynamic OpenDeck(string path,out bool owned){
        string full=System.IO.Path.GetFullPath(path);
        dynamic presentations=ppt.Presentations;
        try{
        // Indexed access avoids retaining Office's hidden COM enumerator.
        for(int i=1;i<=(int)presentations.Count;i++){dynamic deck=presentations[i];bool retained=false;try{if(String.Equals((string)deck.FullName,full,StringComparison.OrdinalIgnoreCase)){
            owned=false;
            if((int)deck.ReadOnly==0)throw new Exception("This presentation is open for editing. Close it in PowerPoint before starting DualSlide: "+full);
            if((int)deck.Windows.Count==0)throw new Exception("This presentation is open invisibly in another session. Close that session first: "+full);
            retained=true;return deck;
        }}finally{if(!retained)ReleaseCom((object)deck);}}
        // Retain a real editor window so PowerPoint cannot auto-close a hidden,
        // dirty presentation (and prompt to save) when its native slideshow exits.
        owned=true;dynamic opened=presentations.Open(full,-1,0,-1);
        // Leave the editor restored until PowerPoint initializes the slideshow renderer.
        return opened;
        }finally{ReleaseCom((object)presentations);}
    }
    dynamic RunShow(dynamic deck,Screen screen,out IntPtr h,out IntPtr canvas){
        deck.Windows[1].WindowState=1;deck.Windows[1].Activate();
        deck.SlideShowSettings.ShowPresenterView=0;
        deck.SlideShowSettings.ShowType=2;deck.SlideShowSettings.ShowScrollbar=0;
        dynamic show=deck.SlideShowSettings.Run();
        canvas=new IntPtr(((ISlideShowHandle)(object)show).HWND);h=Native.GetAncestor(canvas,2);
        var className=new StringBuilder(128);Native.GetClassName(h,className,128);
        if(!Native.IsWindow(h)||!className.ToString().Equals("PPTFrameClass",StringComparison.OrdinalIgnoreCase))throw new Exception("PowerPoint did not create its native slideshow window.");
        SessionLog.Write("Native slideshow "+h+" target "+screen.DeviceName+" "+screen.Bounds);
        Place(h,screen);FitCanvas(canvas,screen);watchdog.Track(masterHandle,h);
        return show;
    }
    Screen SelectedScreen(ComboBox selector){var choice=selector.SelectedItem as DisplayChoice;foreach(var screen in Screen.AllScreens)if(choice!=null&&screen.DeviceName==choice.Device)return screen;throw new Exception("A selected display was disconnected.");}
    void EnsurePlacement(IntPtr h,Screen screen){if(Native.IsIconic(h)||Native.Client(h)==screen.Bounds)return;Rectangle b=screen.Bounds;Native.SetWindowPos(h,IntPtr.Zero,b.X,b.Y,b.Width,b.Height,0x4634);}
    void FitCanvas(IntPtr canvas,Screen screen){
        if(canvas==IntPtr.Zero||Native.IsIconic(Native.GetAncestor(canvas,2))||Native.Client(canvas)==screen.Bounds)return;
        int style=Native.GetWindowLong(canvas,-16);Native.SetWindowLong(canvas,-16,style & ~0x21FF0000);
        Rectangle parent=Native.Client(Native.GetAncestor(canvas,1)),b=screen.Bounds;
        // PowerPoint can leave old non-client padding inside the frame after a
        // DPI change. Fit its actual native MDI slideshow canvas as well.
        Native.SetWindowPos(canvas,IntPtr.Zero,b.X-parent.X,b.Y-parent.Y,b.Width,b.Height,0x4734);
    }
    void Place(IntPtr h,Screen screen){
        // PowerPoint's native windowed slideshow recalculates its canvas across
        // mixed-DPI monitors. Remove only its frame; PowerPoint still renders
        // slides, animations, context menus, and laser controls itself.
        Native.ShowWindow(h,9);int style=Native.GetWindowLong(h,-16);Native.SetWindowLong(h,-16,(style & ~0x21FF0000)|unchecked((int)0x80000000));
        int square=1,border=unchecked((int)0xfffffffe);Native.DwmSetWindowAttribute(h,33,ref square,4);Native.DwmSetWindowAttribute(h,34,ref border,4);
        Rectangle b=screen.Bounds;Native.SetWindowPos(h,IntPtr.Zero,b.X,b.Y,b.Width,b.Height,0x624);
    }
    bool Active(){return running&&!paused&&Native.GetForegroundWindow()==masterHandle;}
    static bool ForwardKey(int key){return key==13||key==32||key==34||key==39||key==40||key==78;}
    bool PresenterAtEnd(){return (int)masterShow.View.Slide.SlideIndex==(int)master.Slides.Count&&(int)masterShow.View.GetClickIndex()>=(int)masterShow.View.GetClickCount();}
    void AdvancePresenter(){if(!PresenterAtEnd())masterShow.View.Next();}
    void QueueStep(int direction){if(!Active())return;navigation.Enqueue(direction);ProcessNavigation();}
    bool Supported(int key){return key==8||key==13||key==27||key==32||(key>=33&&key<=40)||(key>=48&&key<=57)||(key>=96&&key<=105)||key==66||key==87||key==17||key==162||key==163;}
    IntPtr Keyboard(int code,IntPtr msg,IntPtr data){
        if(code>=0){var k=(Native.Key)Marshal.PtrToStructure(data,typeof(Native.Key));int key=(int)k.Code;bool up=msg.ToInt32()==0x101||msg.ToInt32()==0x105;
            IntPtr foreground=Native.GetForegroundWindow();
            if(key==27&&running&&(foreground==masterHandle||foreground==followerHandle)){
                // Consume Escape before PowerPoint can enter its independent close/save path.
                if(!up)BeginInvoke((Action)Stop);return new IntPtr(1);
            }
            bool plain=(Native.GetAsyncKeyState(17)&0x8000)==0&&(Native.GetAsyncKeyState(18)&0x8000)==0&&(Native.GetAsyncKeyState(91)&0x8000)==0&&(Native.GetAsyncKeyState(92)&0x8000)==0;
            if(Active()&&plain){
                if(!up&&((key>=48&&key<=57)||(key>=96&&key<=105)))numberEntry=true;
                if(ForwardKey(key)){
                    // Keep PowerPoint's number+Enter slide selection intact.
                    bool numbered=key==13&&numberEntry;
                    if(!up)numberEntry=false;
                    if(!numbered){if(!up)BeginInvoke((Action)(()=>QueueStep(1)));return new IntPtr(1);}
                }else if(!up&&!((key>=48&&key<=57)||(key>=96&&key<=105)))numberEntry=false;
            }
            if(running&&foreground==followerHandle&&endReached&&plain&&ForwardKey(key))return new IntPtr(1);
            if(!up&&Supported(key)&&Active()&&(Native.GetAsyncKeyState(18)&0x8000)==0&&(Native.GetAsyncKeyState(91)&0x8000)==0&&(Native.GetAsyncKeyState(92)&0x8000)==0)Command(key);
        }return Native.CallNextHookEx(keyHook,code,msg,data);
    }
    IntPtr MouseInput(int code,IntPtr msg,IntPtr data){
        if(code>=0){
            int message=msg.ToInt32();
            if(message==0x202&&blockedEndClick){blockedEndClick=false;return new IntPtr(1);}
            if(message==0x201&&running&&endReached){
                IntPtr foreground=Native.GetForegroundWindow();var mouse=(Native.Mouse)Marshal.PtrToStructure(data,typeof(Native.Mouse));
                if((foreground==masterHandle||foreground==followerHandle)&&Native.Client(foreground).Contains(mouse.Point)&&(Native.GetAsyncKeyState(17)&0x8000)==0){blockedEndClick=true;return new IntPtr(1);}
            }
        }
        if(code>=0&&msg.ToInt32()==0x20A&&Active()){
            var mouse=(Native.Mouse)Marshal.PtrToStructure(data,typeof(Native.Mouse));
            if(Native.Client(masterHandle).Contains(mouse.Point)){
                int delta=unchecked((short)(mouse.Data>>16));
                BeginInvoke((Action)(()=>QueueWheel(delta)));return new IntPtr(1);
            }
        }
        return Native.CallNextHookEx(mouseHook,code,msg,data);
    }
    void QueueWheel(int delta){
        if(!Active())return;
        int now=Environment.TickCount;if(unchecked(now-lastWheel)>750)wheelRemainder=0;lastWheel=now;
        wheelRemainder+=delta;
        while(Math.Abs(wheelRemainder)>=120){int sign=Math.Sign(wheelRemainder);navigation.Enqueue(-sign);wheelRemainder-=sign*120;}
        ProcessNavigation();
    }
    void ProcessNavigation(){
        if(navigating||navigation.Count==0)return;
        if(!Active()){navigation.Clear();return;}
        navigating=true;bool retry=false;
        try{int direction=navigation.Peek();if(direction>0)AdvancePresenter();else masterShow.View.Previous();navigation.Dequeue();MatchNativeState();}
        catch(COMException){retry=true;status.Text="PowerPoint is busy; navigation will retry automatically.";}
        finally{navigating=false;}
        if(!retry&&navigation.Count>0&&running)BeginInvoke((Action)ProcessNavigation);
    }
    void Forward(int msg,IntPtr wp,IntPtr lp){if(!Native.PostMessage(followerHandle,msg,wp,lp))BeginInvoke((Action)(()=>Fail("Input delivery failed; reset the current slide.")));}
    void UpdatePointer(){
        if(!Active()||sourceRect.Width==0||sourceRect.Height==0){pointer.Hide();return;}
        pointer.Configure(sourceRect,targetRect,masterHandle,followerHandle,nativeLaser);
    }
    void ReleaseInputs(){navigation.Clear();wheelRemainder=0;foreach(int key in held)Native.PostMessage(followerHandle,0x101,new IntPtr(key),new IntPtr(unchecked((int)0xC0000001)));held.Clear();if(mouseButtons!=0)Native.PostMessage(followerHandle,0x202,IntPtr.Zero,IntPtr.Zero);mouseButtons=0;}
    void CheckSession(){
        if(checking)return;checking=true;
        try{CheckSessionCore();}finally{checking=false;watchdog.Beat();}
    }
    void CheckSessionCore(){
        if(System.Threading.Interlocked.Exchange(ref recoveryRequested,0)!=0){Stop();return;}

        if(stopping){FinishStop();return;}
        if(!running||navigating)return;
        if(!Native.IsWindow(masterHandle)||!Native.IsWindow(followerHandle)||!Native.IsWindowVisible(masterHandle)||!Native.IsWindowVisible(followerHandle)){Stop();return;}
        if(Screen.AllScreens.Length<2){Stop();return;}
        try{EnsurePlacement(masterHandle,SelectedScreen(masterMonitor));EnsurePlacement(followerHandle,SelectedScreen(followerMonitor));FitCanvas(masterCanvas,SelectedScreen(masterMonitor));FitCanvas(followerCanvas,SelectedScreen(followerMonitor));}catch(Exception){Stop();return;}
        if(!Active()){pointer.Hide();ReleaseInputs();return;}
        sourceRect=Native.Content(Native.Client(masterCanvas),aspect);targetRect=Native.Content(Native.Client(followerCanvas),aspect);
        // Pointer inspection must never prevent slide synchronization.
        try{nativeLaser=(bool)masterShow.View.LaserPointerEnabled;}catch(COMException){}
        try{ProcessNavigation();if(!running)return;MatchNativeState();status.Text=pointer.Error==null?"Synchronized. Presenter controls both slideshows.":"Pointer could not be shown: "+pointer.Error;}
        catch(COMException){status.Text="PowerPoint is transitioning; synchronization will retry automatically.";}
    }
    void Fail(string message){paused=true;pointer.Hide();ReleaseInputs();pause.Text="Resume";status.Text=message;}
    void Stop(){
        if(stopping){FinishStop();return;}
        stopping=true;running=false;endReached=false;numberEntry=blockedEndClick=false;ReleaseInputs();pointer.EndSession();
        if(input!=null){input.Dispose();input=null;}keyHook=mouseHook=IntPtr.Zero;
        start.Enabled=pause.Enabled=stop.Enabled=resync.Enabled=false;status.Text="Closing both slideshows without saving...";
        FinishStop();
    }
    static bool Gone(COMException ex){return ex.ErrorCode==unchecked((int)0x80010108)||ex.ErrorCode==unchecked((int)0x800706BA)||ex.ErrorCode==unchecked((int)0x80048010);}
    bool CloseDeck(dynamic deck,dynamic show,IntPtr handle,bool owned,int showType,int presenterView,int scrollbar,bool restoreSettings){
        if(deck==null)return true;
        try{
            // Suppress the dirty flag BEFORE exiting the show, then again before Close.
            // Saved=true does not write a file. Only tool-owned documents are discarded.
            if(owned)deck.Saved=-1;
            if(show!=null&&Native.IsWindow(handle)&&Native.IsWindowVisible(handle)){
                try{show.View.Exit();}catch(COMException ex){if(!Gone(ex))throw;}
            }
            if(owned){
                string reopenPath=keepDecksForSession?(string)deck.FullName:null;
                deck.Saved=-1;deck.Close();
                if(reopenPath!=null)reopenForEditing.Enqueue(reopenPath);
            }
            else if(restoreSettings){deck.SlideShowSettings.ShowType=showType;deck.SlideShowSettings.ShowPresenterView=presenterView;deck.SlideShowSettings.ShowScrollbar=scrollbar;}
            return true;
        }catch(COMException ex){if(Gone(ex))return true;status.Text="Waiting for PowerPoint to close both shows without saving...";return false;}
    }
    void FinishStop(){
        if(!stopping||cleaning)return;cleaning=true;
        try{
            if(CloseDeck(master,masterShow,masterHandle,ownMaster,masterType,masterPresenter,masterScrollbar,masterConfigured)){ReleaseCom((object)masterShow);ReleaseCom((object)master);master=masterShow=null;masterHandle=masterCanvas=IntPtr.Zero;}
            if(CloseDeck(follower,followerShow,followerHandle,ownFollower,followerType,followerPresenter,followerScrollbar,followerConfigured)){ReleaseCom((object)followerShow);ReleaseCom((object)follower);follower=followerShow=null;followerHandle=followerCanvas=IntPtr.Zero;}
            if(master!=null||follower!=null)return;
            foreach(var guard in sourceGuards)guard.Dispose();sourceGuards.Clear();
            // Reopen only after BOTH read-only decks close and ALL write guards release.
            while(reopenForEditing.Count>0){
                string path=reopenForEditing.Peek();dynamic presentations=null;
                try{
                    presentations=ppt.Presentations;
                    object editable=presentations.Open(path,0,0,-1);
                    reopenForEditing.Dequeue();ReleaseCom(editable);
                }catch(COMException ex){
                    if(ex.ErrorCode==unchecked((int)0x80010001)||ex.ErrorCode==unchecked((int)0x8001010A)){status.Text="Waiting for PowerPoint to reopen presentations for editing...";return;}
                    reopenForEditing.Dequeue();reopenError="Could not reopen for editing: "+path+". "+ex.Message;SessionLog.Write(reopenError);
                }finally{ReleaseCom((object)presentations);}
            }
            if(!ClosePowerPoint())return;
            ownMaster=ownFollower=masterConfigured=followerConfigured=false;
            watchdog.Disarm();stopping=false;paused=false;sourceRect=targetRect=Rectangle.Empty;
            start.Enabled=true;foreach(var control in configuration)control.Enabled=true;pause.Text="Pause";status.Text=reopenError??(keepDecksForSession?"Stopped. Presentations left open for editing. DualSlide saved no files.":"Stopped. Original files unchanged.");keepDecksForSession=false;
            if(closeAfterStop){closeAfterStop=false;BeginInvoke((Action)(()=>{if(!IsDisposed)Close();}));}
        }finally{cleaning=false;}
    }
    bool ClosePowerPoint(){
        if(ppt==null){ownPowerPoint=false;return true;}
        try{
            if(ownPowerPoint){
                dynamic presentations=ppt.Presentations;
                try{if((int)presentations.Count==0)ppt.Quit();}
                finally{ReleaseCom((object)presentations);}
            }
        }catch(COMException ex){if(!Gone(ex)){status.Text="Waiting for PowerPoint to quit...";return false;}}
        ReleaseCom((object)ppt);ppt=null;ownPowerPoint=false;
        // Dynamic Office property chains create temporary RCWs. Drain those after
        // all shows and decks close so they cannot keep our Office server alive.
        GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();GC.WaitForPendingFinalizers();
        return true;
    }
}

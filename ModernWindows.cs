using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

static class FluentDrawing {
    public static System.Drawing.Drawing2D.GraphicsPath Rounded(RectangleF r,float radius){var p=new System.Drawing.Drawing2D.GraphicsPath();float d=radius*2;p.AddArc(r.X,r.Y,d,d,180,90);p.AddArc(r.Right-d,r.Y,d,d,270,90);p.AddArc(r.Right-d,r.Bottom-d,d,d,0,90);p.AddArc(r.X,r.Bottom-d,d,d,90,90);p.CloseFigure();return p;}
    public static Color Blend(Color a,Color b,double amount){return Color.FromArgb((int)(a.R+(b.R-a.R)*amount),(int)(a.G+(b.G-a.G)*amount),(int)(a.B+(b.B-a.B)*amount));}
}
sealed class ModernCard:Panel {
    public Color Surface=Color.White,BorderColor=Color.FromArgb(225,225,225);
    public ModernCard(){SetStyle(ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer,true);}
    protected override void OnPaintBackground(PaintEventArgs e){e.Graphics.Clear(Parent==null?BackColor:Parent.BackColor);e.Graphics.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;using(var path=FluentDrawing.Rounded(new RectangleF(.5f,.5f,Width-1,Height-1),8)){using(var brush=new SolidBrush(Surface))e.Graphics.FillPath(brush,path);using(var pen=new Pen(BorderColor))e.Graphics.DrawPath(pen,path);}}
}
sealed class ModernButton:Button {
    public Color FillColor=Color.White,BorderColor=Color.LightGray,Accent=Color.Blue,DisabledFill=Color.Gainsboro,DisabledText=Color.Gray;
    public bool Primary;bool hovered,pressed;
    public ModernButton(){SetStyle(ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.ResizeRedraw,true);FlatStyle=FlatStyle.Flat;FlatAppearance.BorderSize=0;UseVisualStyleBackColor=false;Cursor=Cursors.Hand;}
    // Clear every corner with the parent's actual surface; no native Button border is painted.
    protected override void OnPaintBackground(PaintEventArgs e){e.Graphics.Clear(Parent==null?BackColor:Parent.BackColor);}
    protected override void OnPaint(PaintEventArgs e){OnPaintBackground(e);e.Graphics.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        Color fill=Enabled?FillColor:DisabledFill;if(Enabled&&hovered)fill=FluentDrawing.Blend(fill,ForeColor,pressed?.12:.055);
        using(var path=FluentDrawing.Rounded(new RectangleF(.75f,.75f,Width-1.5f,Height-1.5f),5)){using(var brush=new SolidBrush(fill))e.Graphics.FillPath(brush,path);if(!Primary)using(var pen=new Pen(BorderColor))e.Graphics.DrawPath(pen,path);}
        TextRenderer.DrawText(e.Graphics,Text,Font,ClientRectangle,Enabled?ForeColor:DisabledText,TextFormatFlags.HorizontalCenter|TextFormatFlags.VerticalCenter|TextFormatFlags.SingleLine);
        if(Focused&&ShowFocusCues)using(var path=FluentDrawing.Rounded(new RectangleF(3,3,Width-6,Height-6),3))using(var pen=new Pen(Primary?ForeColor:Accent,1.5f))e.Graphics.DrawPath(pen,path);
    }
    protected override void OnMouseEnter(EventArgs e){hovered=true;Invalidate();base.OnMouseEnter(e);}protected override void OnMouseLeave(EventArgs e){hovered=pressed=false;Invalidate();base.OnMouseLeave(e);}
    protected override void OnMouseDown(MouseEventArgs e){pressed=true;Invalidate();base.OnMouseDown(e);}protected override void OnMouseUp(MouseEventArgs e){pressed=false;Invalidate();base.OnMouseUp(e);}
    protected override void OnEnabledChanged(EventArgs e){Invalidate();base.OnEnabledChanged(e);}
}
sealed class ModernComboBox:ComboBox {
    public Color BorderColor=Color.LightGray,Accent=Color.Blue;
    public ModernComboBox(){FlatStyle=FlatStyle.Flat;DrawMode=DrawMode.OwnerDrawFixed;ItemHeight=28;}
    protected override void OnDrawItem(DrawItemEventArgs e){if(e.Index<0)return;bool selected=(e.State&DrawItemState.Selected)!=0;Color background=selected?FluentDrawing.Blend(BackColor,Accent,.18):BackColor;using(var brush=new SolidBrush(background))e.Graphics.FillRectangle(brush,e.Bounds);var text=e.Bounds;text.Inflate(-8,0);TextRenderer.DrawText(e.Graphics,Items[e.Index].ToString(),Font,text,ForeColor,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.EndEllipsis|TextFormatFlags.SingleLine);}
    protected override void WndProc(ref Message message){base.WndProc(ref message);if(message.Msg==0xF&&IsHandleCreated){using(var g=Graphics.FromHwnd(Handle))PaintFrame(g);}else if(message.Msg==0x318&&message.WParam!=IntPtr.Zero){using(var g=Graphics.FromHdc(message.WParam))PaintFrame(g);}}
    void PaintFrame(Graphics g){g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;int arrow=Math.Max(28,Height);using(var brush=new SolidBrush(BackColor))g.FillRectangle(brush,Width-arrow,0,arrow,Height);
        var bounds=new RectangleF(.75f,.75f,Width-1.5f,Height-1.5f);using(var path=FluentDrawing.Rounded(bounds,4)){using(var outside=new Region(ClientRectangle)){outside.Exclude(path);using(var background=new SolidBrush(Parent==null?BackColor:Parent.BackColor))g.FillRegion(background,outside);}using(var pen=new Pen(Focused?Accent:BorderColor))g.DrawPath(pen,path);}
        float x=Width-arrow/2f,y=Height/2f;using(var pen=new Pen(Enabled?ForeColor:SystemColors.GrayText,1.5f)){g.DrawLines(pen,new[]{new PointF(x-4,y-2),new PointF(x,y+2),new PointF(x+4,y-2)});}
    }
    protected override void OnGotFocus(EventArgs e){base.OnGotFocus(e);Invalidate();}protected override void OnLostFocus(EventArgs e){base.OnLostFocus(e);Invalidate();}
}
static class ModernWindows {
    [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr h,int name,ref int value,int size);
    [DllImport("dwmapi.dll")] static extern int DwmGetColorizationColor(out uint color,out bool opaque);
    [DllImport("uxtheme.dll",CharSet=CharSet.Unicode)] static extern int SetWindowTheme(IntPtr h,string theme,string sub);
    public static bool Dark{get{try{return Convert.ToInt32(Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize","AppsUseLightTheme",1))==0;}catch{return false;}}}
    public static Color Accent{get{try{uint color;bool opaque;if(DwmGetColorizationColor(out color,out opaque)==0)return Color.FromArgb((int)((color>>16)&255),(int)((color>>8)&255),(int)(color&255));}catch{}return SystemColors.Highlight;}}
    static double Linear(byte x){double n=x/255.0;return n<=.04045?n/12.92:Math.Pow((n+.055)/1.055,2.4);}
    public static Color AccentText(Color accent){double luminance=.2126*Linear(accent.R)+.7152*Linear(accent.G)+.0722*Linear(accent.B);return (1.05/(luminance+.05))>((luminance+.05)/.05)?Color.White:Color.Black;}
    public static void Apply(Form form,bool? preview=null){bool dark=(preview??Dark)&&!SystemInformation.HighContrast;
        Color bg=dark?Color.FromArgb(32,32,32):Color.FromArgb(243,243,243),card=dark?Color.FromArgb(39,39,39):Color.FromArgb(251,251,251),surface=dark?Color.FromArgb(50,50,50):Color.White,fg=dark?Color.FromArgb(245,245,245):Color.FromArgb(26,26,26),muted=dark?Color.FromArgb(190,190,190):Color.FromArgb(96,96,96),border=dark?Color.FromArgb(65,65,65):Color.FromArgb(220,220,220),accent=Accent;
        if(SystemInformation.HighContrast){bg=card=SystemColors.Window;surface=SystemColors.Control;fg=muted=SystemColors.WindowText;border=SystemColors.WindowText;accent=SystemColors.Highlight;}
        form.BackColor=bg;form.ForeColor=fg;int enabled=dark?1:0,rounded=2,mica=2;DwmSetWindowAttribute(form.Handle,20,ref enabled,4);DwmSetWindowAttribute(form.Handle,33,ref rounded,4);DwmSetWindowAttribute(form.Handle,38,ref mica,4);
        Style(form,bg,card,surface,fg,muted,border,accent,dark);form.Invalidate(true);
    }
    static void Style(Control parent,Color background,Color card,Color surface,Color fg,Color muted,Color border,Color accent,bool dark){foreach(Control control in parent.Controls){control.ForeColor=Object.Equals(control.Tag,"muted")?muted:fg;control.BackColor=background;
        var panel=control as ModernCard;if(panel!=null){panel.Surface=card;panel.BorderColor=border;control.BackColor=card;Style(control,card,card,surface,fg,muted,border,accent,dark);continue;}
        var button=control as ModernButton;if(button!=null){button.FlatAppearance.BorderSize=0;button.Primary=button.Text=="Start presentation";button.FillColor=button.Primary?accent:surface;button.ForeColor=button.Primary?AccentText(accent):fg;button.BorderColor=border;button.Accent=accent;button.DisabledFill=FluentDrawing.Blend(background,fg,.045);button.DisabledText=muted;}
        var combo=control as ModernComboBox;if(combo!=null){combo.BorderColor=border;combo.Accent=accent;}
        if(control is ComboBox){control.BackColor=surface;SetWindowTheme(control.Handle,dark?"DarkMode_Explorer":"Explorer",null);}
        var link=control as LinkLabel;if(link!=null){link.LinkColor=accent;link.ActiveLinkColor=accent;link.VisitedLinkColor=accent;}
        if(control is CheckBox)SetWindowTheme(control.Handle,dark?"DarkMode_Explorer":"Explorer",null);
    }}
}

sealed class DisplayChoice {
    public string Device,Name,Number;
    public override string ToString(){return Number+" \u00B7 "+Name;}
}

static class Displays {
    [StructLayout(LayoutKind.Sequential)] struct Luid{public uint Low;public int High;}
    [StructLayout(LayoutKind.Sequential)] struct Rational{public uint Numerator,Denominator;}
    [StructLayout(LayoutKind.Sequential)] struct Source{public Luid Adapter;public uint Id,Mode,Status;}
    [StructLayout(LayoutKind.Sequential)] struct Target{public Luid Adapter;public uint Id,Mode,Technology,Rotation,Scaling;public Rational Refresh;public uint Scan,Available,Status;}
    [StructLayout(LayoutKind.Sequential)] struct Path{public Source Source;public Target Target;public uint Flags;}
    [StructLayout(LayoutKind.Sequential)] struct Header{public uint Type,Size;public Luid Adapter;public uint Id;}
    [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] struct SourceName{public Header Header;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=32)]public string Name;}
    [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] struct TargetName{public Header Header;public uint Flags,Technology;public ushort Manufacturer,Product;public uint Connector;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=64)]public string Friendly;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=128)]public string DevicePath;}
    [DllImport("user32.dll")] static extern int GetDisplayConfigBufferSizes(uint flags,out uint paths,out uint modes);
    [DllImport("user32.dll")] static extern int QueryDisplayConfig(uint flags,ref uint count,[Out]Path[] paths,ref uint modeCount,[Out]byte[] modes,IntPtr topology);
    [DllImport("user32.dll",EntryPoint="DisplayConfigGetDeviceInfo")] static extern int DeviceInfo(ref SourceName name);
    [DllImport("user32.dll",EntryPoint="DisplayConfigGetDeviceInfo")] static extern int DeviceInfo(ref TargetName name);
    public static List<DisplayChoice> Read(){var names=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        for(int retry=0;retry<3;retry++){uint pc,mc;if(GetDisplayConfigBufferSizes(2,out pc,out mc)!=0)break;var paths=new Path[pc];var modes=new byte[mc*64];int result=QueryDisplayConfig(2,ref pc,paths,ref mc,modes,IntPtr.Zero);if(result==122)continue;if(result!=0)break;
            for(int i=0;i<pc;i++){var source=new SourceName{Header=new Header{Type=1,Size=(uint)Marshal.SizeOf(typeof(SourceName)),Adapter=paths[i].Source.Adapter,Id=paths[i].Source.Id}};var target=new TargetName{Header=new Header{Type=2,Size=(uint)Marshal.SizeOf(typeof(TargetName)),Adapter=paths[i].Target.Adapter,Id=paths[i].Target.Id}};
                if(DeviceInfo(ref source)==0&&DeviceInfo(ref target)==0)names[source.Name]=!String.IsNullOrWhiteSpace(target.Friendly)?target.Friendly:((target.Technology==6||target.Technology==11||target.Technology==0x80000000)?"Internal display":"Display");
            }break;
        }
        var choices=new List<DisplayChoice>();foreach(var screen in Screen.AllScreens){string name;if(!names.TryGetValue(screen.DeviceName,out name))name=screen.Primary?"Built-in / primary display":"Display";string number=System.Text.RegularExpressions.Regex.Match(screen.DeviceName,@"\d+$").Value;choices.Add(new DisplayChoice{Device=screen.DeviceName,Name=name,Number=number});}return choices;
    }
    public static void Identify(){var windows=new List<Form>();foreach(var choice in Read()){Screen screen=Array.Find(Screen.AllScreens,s=>s.DeviceName==choice.Device);if(screen==null)continue;var popup=new Identification(choice.Number,choice.Name);popup.StartPosition=FormStartPosition.Manual;popup.Bounds=new Rectangle(screen.Bounds.X+(screen.Bounds.Width-360)/2,screen.Bounds.Y+(screen.Bounds.Height-240)/2,360,240);popup.Show();windows.Add(popup);}
        var timer=new Timer{Interval=3000};timer.Tick+=(s,e)=>{timer.Stop();foreach(var window in windows)window.Dispose();timer.Dispose();};timer.Start();
    }
    sealed class Identification:Form {
        readonly string number,name;
        public Identification(string number,string name){this.number=number;this.name=name;FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;TopMost=true;BackColor=Color.FromArgb(24,24,24);DoubleBuffered=true;}
        protected override bool ShowWithoutActivation{get{return true;}}
        protected override CreateParams CreateParams{get{var p=base.CreateParams;p.ExStyle|=0x08000080;return p;}}
        protected override void OnPaint(PaintEventArgs e){e.Graphics.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;using(var font=new Font("Segoe UI Variable Display",90,FontStyle.Bold))using(var format=new StringFormat{Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center})e.Graphics.DrawString(number,font,Brushes.White,new RectangleF(0,0,Width,185),format);using(var font=new Font("Segoe UI",12))using(var format=new StringFormat{Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center})e.Graphics.DrawString(name,font,Brushes.White,new RectangleF(10,185,Width-20,45),format);}
    }
}

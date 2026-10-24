using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using Point=System.Drawing.Point;
class PointerRegressionTest {
 [DllImport("user32.dll")]static extern int GetGuiResources(IntPtr process,int flags);
 static void Assert(bool value,string message){if(!value)throw new Exception(message);}
 [STAThread]static void Main(){try{
  Point hotspot;var cursor=Cursors.Hand.Handle;var source=PointerForm.CursorArtwork(cursor,out hotspot);
  Assert(source.PixelWidth>0&&source.PixelHeight>0&&source.IsFrozen,"Native cursor conversion failed");
  Assert(hotspot.X>=0&&hotspot.X<source.PixelWidth&&hotspot.Y>=0&&hotspot.Y<source.PixelHeight,"Invalid cursor hotspot");
  var pixels=new byte[source.PixelWidth*source.PixelHeight*4];source.CopyPixels(pixels,source.PixelWidth*4,0);
  bool transparent=false,opaque=false;for(int i=3;i<pixels.Length;i+=4){transparent|=pixels[i]==0;opaque|=pixels[i]==255;}Assert(transparent&&opaque,"Cursor transparency/opaque outline lost");
  int before=GetGuiResources(Process.GetCurrentProcess().Handle,0);for(int i=0;i<1000;i++)PointerForm.CursorArtwork(cursor,out hotspot);int after=GetGuiResources(Process.GetCurrentProcess().Handle,0);Assert(after-before<5,"Cursor copying leaked GDI handles");
  var laser=PointerForm.LaserArtwork(255);laser.Measure(new System.Windows.Size(32,32));laser.Arrange(new Rect(0,0,32,32));var bitmap=new RenderTargetBitmap(32,32,96,96,PixelFormats.Pbgra32);bitmap.Render(laser);pixels=new byte[32*32*4];bitmap.CopyPixels(pixels,128,0);
  int center=pixels[(16*32+16)*4+3],ring=pixels[(16*32+19)*4+3],edge=pixels[3];Assert(ring>center&&center>edge,"Laser must have a ring, translucent center, and fading halo");
  Console.WriteLine("PASS: native cursor alpha/hotspot, 1000 conversions without GDI leaks, laser ring/center/halo.");
 }catch(Exception ex){Console.WriteLine("FAIL: "+ex);Environment.ExitCode=1;}}
}

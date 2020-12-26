using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExToolsForExcelTest
{
    class WindowCapture
    {
        public static Bitmap CaptureActiveWindow()
        {
            Rect rect;
            // アクティブウィンドウを取得
            IntPtr activeWindow = GetForegroundWindow();
            GetWindowRect(activeWindow, out rect);
            Rectangle rectangle = new Rectangle(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);


            Bitmap bitmap = new Bitmap(rectangle.Width, rectangle.Height);
            Graphics graphics = Graphics.FromImage(bitmap);


            // 画面をコピー
            graphics.CopyFromScreen(new Point(rectangle.X, rectangle.Y), new Point(0, 0), rectangle.Size);

            return bitmap;
        }
        public static bool SaveActiveWindowCapture(string path)
        {
            return SaveActiveWindowCapture(CaptureActiveWindow(), path);
        }
        public static bool SaveActiveWindowCapture(Bitmap bmp,string path)
        {
            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            return true;
        }


        [DllImport("user32.Dll")]
        static extern int GetWindowRect(IntPtr hWnd, out Rect rect);
        [DllImport("user32.dll")]
        extern static IntPtr GetForegroundWindow();

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
    }
}

/*
* ==============================================================================
*
* File name: Win32
* Description: IE common library, Version 1.0
* Host Access Class Library
*
* Version: 1.0
* Created: 19/11/2018 21:36:16
*
* Author: Haley X L Zhang
* Company: Chinasoft International
*
* ==============================================================================
*/
using Microsoft.VisualStudio.CommandBars;
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace IEOperateLib.Common
{
    public class Win32
    {
        #region Windows - user32.dll


        public enum hWndInsertAfter
        {
            /// <summary>
            /// Places the window at the bottom of Z order. 
            /// If yhe hWnd parameter identifies a topmost window, 
            /// the window loses its topmost status and is placed at the bottom of all other windows.
            /// </summary>
            HWND_BOTTON = 1,
            /// <summary>
            /// Places the window above all non-topmost windows(that is, behind all topmost windows). 
            /// This flag has noeffect if the window is already a non-topmost window.
            /// </summary>
            HWND_NOTTOPMOST = -2,
            /// <summary>
            /// Places the window at the top of the Z order.
            /// </summary>
            HWND_TOP = 0,
            /// <summary>
            /// Places the window above all non-topmost windows.
            /// The window maintains its topmost position even when it is deactivated.
            /// </summary>
            HWND_TOPMOST = -1
        }

        public struct Rect
        {
            public int left;
            public int top;
            public int right;
            public int bottom;

            public int Width
            {
                get
                {
                    return right - left;
                }
            }

            public int Height
            {
                get
                {
                    return bottom - top;
                }
            }
        }

        public enum GetWindow_Cmd : uint
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }

        #region  SetWindowPos  uFlags
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        public static readonly IntPtr HWND_TOP = new IntPtr(0);
        public const UInt32 SWP_NOSIZE = 0x0001;
        public const UInt32 SWP_NOMOVE = 0x0002;
        public const UInt32 SWP_NOZORDER = 0x0004;
        public const UInt32 SWP_NOREDRAW = 0x0008;
        public const UInt32 SWP_NOACTIVATE = 0x0010;
        public const UInt32 SWP_FRAMECHANGED = 0x0020;
        public const UInt32 SWP_SHOWWINDOW = 0x0040;
        public const UInt32 SWP_HIDEWINDOW = 0x0080;
        public const UInt32 SWP_NOCOPYBITS = 0x0100;
        public const UInt32 SWP_NOOWNERZORDER = 0x0200;
        public const UInt32 SWP_NOSENDCHANGING = 0x0400;
        public const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int BM_CLICK = 0xF5;
        public const UInt32 CHILDID_SELF = 0;
        public const UInt32 CHILDID_1 = 1;
        public const UInt32 OBJID_CLIENT = 0xFFFFFFC;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYUP = 0x101;
        public const int WM_SETTEXT = 0x000C;
        public const int VK_F6 = 117;
        public const int WM_MOUSEWHEEL = 0x20A;
        //双击鼠标中键   
        public const int WM_MBUTTONDBLCLK = 0x209;
        //释放鼠标中键   
        public const int WM_MBUTTONUP = 0x208;
        //移动鼠标时发生，同WM_MOUSEFIRST   
        public const int WM_MOUSEMOVE = 0x200;

        //释放鼠标左键   
        public const int WM_LBUTTONUP = 0x202;
        //双击鼠标左键   
        public const int WM_LBUTTONDBLCLK = 0x203;
        //按下鼠标右键   
        public const int WM_RBUTTONDOWN = 0x204;
        //释放鼠标右键   
        public const int WM_RBUTTONUP = 0x205;
        //双击鼠标右键   
        public const int WM_RBUTTONDBLCLK = 0x206;
        //按下鼠标中键   
        public const int WM_MBUTTONDOWN = 0x207;

        public const int WM_USER = 0x0400;
        public const int MK_LBUTTON = 0x0001;
        public const int MK_RBUTTON = 0x0002;
        public const int MK_SHIFT = 0x0004;
        public const int MK_CONTROL = 0x0008;
        public const int MK_MBUTTON = 0x0010;
        public const int VK_TAB = 9;
        public const int VK_ENTER = 13;



        /// <summary>
        /// 
        /// </summary>
        /// <param name="hWnd">handl of window</param>
        /// <param name="hWndInsertAfter">placement-order handle</param>
        /// <param name="X">position</param>
        /// <param name="Y">position</param>
        /// <param name="cx">width</param>
        /// <param name="cy">height</param>
        /// <param name="uFlags">window-position flags</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        #endregion
        /// <summary>
        /// //获取窗体句柄，hwnd为源窗口句柄
            /*wCmd指定结果窗口与源窗口的关系，它们建立在下述常数基础上：
                  GW_CHILD
                  寻找源窗口的第一个子窗口
                  GW_HWNDFIRST
                  为一个源子窗口寻找第一个兄弟（同级）窗口，或寻找第一个顶级窗口
                  GW_HWNDLAST
                  为一个源子窗口寻找最后一个兄弟（同级）窗口，或寻找最后一个顶级窗口
                  GW_HWNDNEXT
                  为源窗口寻找下一个兄弟窗口
                  GW_HWNDPREV
                  为源窗口寻找前一个兄弟窗口
                  GW_OWNER
                  寻找窗口的所有者
             */
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="wCmd"></param>
        /// <returns></returns>
        [DllImport("User32.dll")]
        public static extern Int32 GetWindow(IntPtr hWnd, Int32 wCmd);
        [DllImport("User32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);



        [DllImport("user32.dll")]
        public static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpString, int nMaxCount);


        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr window);
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);
        [DllImport("User32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("User32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
        [DllImport("user32.dll")]
        public static extern int PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);


        /// <summary>
        /// 这个函数有四个参数，第一个是窗口句柄，窗口可以是任何类型的屏幕对象；
        /// 第二个是用于区别其他消息的常量值；
        /// 第三个通常是一个与消息有关的常量值，也可能是窗口或控件的句柄，
        /// 第三个参数是可选参数，有的消息要，有的不需要，比如单击就不需要这个参数，
        /// 别的消息，比如鼠标移动的可能需要在这里加上一些鼠标的参数；
        /// 第四个通常是一个指向内存中数据的指针。
        /// 在C#中消息需要定义成windows系统中的原始的16进制数字，
        /// 比如 const int WM_Lbutton = 0x201; 
        /// </summary>
        /// <param name="hwnd">窗口句柄，窗口可以是任何类型的屏幕对象</param>
        /// <param name="wMsg">区别其他消息的常量值</param>
        /// <param name="wParam">通常是一个与消息有关的常量值，也可能是窗口或控件的句柄
        /// 参数是可选参数，有的消息要，有的不需要，比如单击就不需要这个参数
        /// </param>
        /// <param name="lParam">通常是一个指向内存中数据的指针。在C#中消息需要定义成windows系统中的原始的16进制数字</param>
        /// <returns></returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, string lParam);

        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        #endregion Windows

        #region gdi32
        public const int SRCCOPY = 13369376;

        [DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
        public static extern IntPtr DeleteDC(IntPtr hDc);

        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        public static extern IntPtr DeleteObject(IntPtr hDc);

        [DllImport("gdi32.dll", EntryPoint = "BitBlt")]
        public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, int RasterOp);

        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr bmp);
        #endregion


        [DllImport("Oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, int dwObjectID, ref Guid refID, ref IAccessible ppvObject);
        [DllImport("Oleacc.dll")]
        public static extern int AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);
    }
}

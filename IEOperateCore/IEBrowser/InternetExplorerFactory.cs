using IEOperateLib.Common;
using SHDocVw;
using System;

namespace IEOperateLib.IEBrowser
{
    internal class InternetExplorerFactory
    {
        private InternetExplorerFactory() { }
        private static InternetExplorer IE = null;
        public static InternetExplorer GetInternetExplorer()
        {
            IE = new InternetExplorer();
            IE.Visible = true;
            return IE;
        }
        public static void CloseInternetExplorer()
        {
            if (IE != null)
            {
               
                IE.Stop();
                IE.Quit();
                Win32.SendMessage((IntPtr)IE.HWND, 16, new IntPtr(0), new IntPtr(0));
                
                IE = null;
            }
        }
        public static InternetExplorer GetInternetExplorer(string url)
        {
           
            bool ieOpened = false;
            //使用Microsoft Internet Controls取得所有的已经打开的IE(以Tab计算)
            SHDocVw.ShellWindows ieTabs = new SHDocVw.ShellWindows();
            //每个一个Tab都可以操作，每个Tab对应Com Object的SHDocVw.InternetExplorer
            foreach (SHDocVw.InternetExplorer ieTab in ieTabs)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ieTab.FullName).ToLower();
                if (filename.Equals("iexplore") && ieTab.LocationURL.Equals(url))
                {
                    new TabActivator((IntPtr)ieTab.HWND).ActivateByTabsUrl(ieTab.LocationURL);
                    IE = ieTab;
                    IE.Visible = true;
                    ieOpened = true;
                    break;
                }
            }
            if (!ieOpened)
            {
                IE = new InternetExplorer();
                IE.Visible = true;
                IE.Navigate(url);
            }
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_TOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_NOTTOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);
            return IE;
        }

        public static bool IsMatchIEWindow(string url)
        {

            bool ieOpened = false;
            //使用Microsoft Internet Controls取得所有的已经打开的IE(以Tab计算)
            SHDocVw.ShellWindows ieTabs = new SHDocVw.ShellWindows();
            //每个一个Tab都可以操作，每个Tab对应Com Object的SHDocVw.InternetExplorer
            foreach (SHDocVw.InternetExplorer ieTab in ieTabs)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ieTab.FullName).ToLower();
                if (filename.Equals("iexplore") && ieTab.LocationURL.Equals(url))
                {
                    new TabActivator((IntPtr)ieTab.HWND).ActivateByTabsUrl(ieTab.LocationURL);
                    IE = ieTab;
                    IE.Visible = true;
                    ieOpened = true;
                    break;
                }
            }
            if (!ieOpened)
            {
                IE = new InternetExplorer();
                IE.Visible = true;
                IE.Navigate(url);
            }
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_TOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_NOTTOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);
            
            return ieOpened;
        }

    }
}

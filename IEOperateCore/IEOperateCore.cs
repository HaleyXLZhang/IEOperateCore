using IEOperateCore.IEBrowser;
using IEOperateCore.Interface;
using SHDocVw;
using mshtml;
using System.Threading;
using System;
using IEOperateCore.Common;
using System.Collections.Generic;
using System.Windows.Forms;

namespace IEOperateCore
{
    public class IEOperateCore : IOperateDom
    {
        public IntPtr HWND = new IntPtr(0);
        public InternetExplorer IE;
        private HTMLDocumentClass dom;

        public IEOperateCore()
        {
            IE = InternetExplorerFactory.GetInternetExplorer();
            HWND = new IntPtr(IE.HWND);

        }
        public IEOperateCore(string url)
        {
            IE = InternetExplorerFactory.GetInternetExplorer(url);
            HWND = new IntPtr(IE.HWND);
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                try
                {
                    dom = (HTMLDocumentClass)IE.Document;
                }
                catch (Exception e)
                {
                    Thread.Sleep(1000);
                    continue;
                }
                if (dom.readyState.Equals("complete")) break;

                if (loopCount > 2000)
                    throw new Exception("open " + url + " timeout!");
                Thread.Sleep(500);
                loopCount++;
            }
        }
        public void SetIETabActivate()
        {
            new TabActivator((IntPtr)IE.HWND).ActivateByTabsUrl(IE.LocationURL);
        }
        public void CloseInternetExplorer()
        {

            if (dom != null)
                dom.close();
            if (IE != null)
                IE.Quit();
        }

        public void Dispose()
        {
            CloseInternetExplorer();
        }
        public bool InternetExplorerWindowIsReady(string url)
        {
            SHDocVw.ShellWindows ieTabs = new SHDocVw.ShellWindows();

            foreach (SHDocVw.InternetExplorer ieTab in ieTabs)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ieTab.FullName).ToLower();

                if (filename.Equals("iexplore") && ieTab.LocationURL.Equals(url))
                {
                    return !ieTab.Busy;
                }
            }
            return false;
        }

        public HTMLDocumentClass GetDom()
        {
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                if (loopCount > 2000)
                    throw new Exception("Get dom timeout!");
                Thread.Sleep(500);
                loopCount++;
            }

            dom = (HTMLDocumentClass)IE.Document;


            return dom;
        }

        public T GetInputElementByID<T>(string id)
        {


            GetDom();
            return (T)dom.getElementById(id);
        }

        public void SenKey(string keyBoard)
        {
            SendKeys.SendWait(keyBoard);
        }

        public void GoBack()
        {
            IE.GoBack();
        }

        public void OpenInternetExplorer(string url)
        {
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_TOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_NOTTOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            IE.Navigate(url);
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                try
                {
                    dom = (HTMLDocumentClass)IE.Document;
                }
                catch (Exception e)
                {
                    Thread.Sleep(1000);
                    continue;
                }
                if (dom.readyState.Equals("complete")) break;
                if (loopCount > 2000)
                    throw new Exception("Get " + url + " timeout!");
                Thread.Sleep(500);
                loopCount++;
            }
        }

        public void Refresh()
        {

            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                Thread.Sleep(500);
            }
            IE.Refresh();
        }

        public void SetInternetExplorerWindowPosition(int top, int left)
        {
            IE.Top = top;
            IE.Left = left;
        }

        public void SetInternetExplorerWindowSize(int height, int width)
        {
            IE.Height = height;
            IE.Width = width;
        }

        public IList<T> getElementByTagName<T>(string tagName)
        {
            IList<T> getCollection = new List<T>();

            GetDom();

            IHTMLElementCollection collection = dom.getElementsByTagName(tagName);

            foreach (IHTMLElement elem in collection)
            {
                if (null != elem)
                    getCollection.Add((T)elem);
            }
            return getCollection;
        }

        public IList<T> getElementByName<T>(string name)
        {
            IList<T> getCollection = new List<T>();

            GetDom();

            IHTMLElementCollection collection = dom.getElementsByName(name);

            foreach (IHTMLElement elem in collection)
            {
                if (null != elem)
                    getCollection.Add((T)elem);
            }
            return getCollection;
        }


        public IntPtr FindWindow(string lpClassName, string lpWindowName)
        {
            return Win32.FindWindow(lpClassName, lpWindowName);
        }
        public List<HtmlElement> GetHtmlElementByTagName(IHTMLElement pieceDom, string tagName)
        {

            List<HtmlElement> getCollection = new List<HtmlElement>();

            //HtmlElementCollection findCollection = pieceDom.innerText(tagName);


            //foreach (HtmlElement element in findCollection)
            //{
            //    if (null != element)
            //        getCollection.Add(element);
            //}
            return getCollection;
        }


        /// <summary>
        /// 此方法是查找iframe下的a标签,查找规则是根据提供的searchKey查找指定的属性内容，默认直返回第一个被查找到的元素
        /// </summary>
        /// <param name="iframe"></param>
        /// <param name="searchKey"></param>
        /// <param name="isFindByTitle"></param>
        /// <param name="isFindByClassName"></param>
        /// <returns></returns>
        public HTMLAnchorElementClass Get_Alink_From_IFrame(HTMLFrameElementClass iframe, string searchKey, bool isFindByTitle = true, bool isFindByClassName = false)
        {
            HTMLAnchorElementClass aLink = null;
            IHTMLElementCollection aLinkElementCollection;
            IHTMLDocument2 doc;
            HTMLBodyClass bodyClass;
            doc = iframe.contentWindow.document;
            bodyClass = (HTMLBodyClass)doc.body;
            aLinkElementCollection = bodyClass.getElementsByTagName("a");
            foreach (HTMLAnchorElementClass item in aLinkElementCollection)
            {
                if (isFindByTitle)
                {
                    if (null != item.title && item.title.Equals(searchKey))
                    {
                        aLink = item;
                        break;
                    }
                }
                else if (isFindByClassName)
                {
                    if (null != item.className && item.className.Equals(searchKey))
                    {
                        aLink = item;
                        break;
                    }
                }
            }
            return aLink;
        }

    }
}

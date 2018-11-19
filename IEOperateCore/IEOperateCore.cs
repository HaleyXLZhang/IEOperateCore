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
        private InternetExplorer ie;
        private HTMLDocumentClass dom;
        public IEOperateCore()
        {
            ie = InternetExplorerFactory.GetInternetExplorer();

        }
        public IEOperateCore(string url)
        {
            ie = InternetExplorerFactory.GetInternetExplorer(url);
            OpenInternetExplorer(url);


        }
        public void SetIETabActivate()
        {
            new TabActivator((IntPtr)ie.HWND).ActivateByTabsUrl(ie.LocationURL);
        }
        public void CloseInternetExplorer()
        {

            if (dom != null)
                dom.close();
            if (ie != null)
                ie.Quit();
        }

        public void Dispose()
        {
            CloseInternetExplorer();
        }

        public HTMLDocumentClass GetDom()
        {

            while (ie.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                Thread.Sleep(500);
            }

            dom = (HTMLDocumentClass)ie.Document;
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
            ie.GoBack();
        }

        public void OpenInternetExplorer(string url)
        {
            Win32.SetWindowPos(new IntPtr(ie.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_TOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            Win32.SetWindowPos(new IntPtr(ie.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_NOTTOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            ie.Navigate(url);
        }

        public void Refresh()
        {

            while (ie.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                Thread.Sleep(500);
            }
            ie.Refresh();
        }

        public void SetInternetExplorerWindowPosition(int top, int left)
        {
            ie.Top = top;
            ie.Left = left;
        }

        public void SetInternetExplorerWindowSize(int height, int width)
        {
            ie.Height = height;
            ie.Width = width;
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
    }
}

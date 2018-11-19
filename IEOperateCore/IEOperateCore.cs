using IEOperateCore.IEBrowser;
using IEOperateCore.Interface;
using SHDocVw;
using mshtml;
using System.Threading;
using System;
using IEOperateCore.Common;

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


        public T GetInputElementByID<T>(HTMLDocumentClass dom, string id)
        {
            return (T)dom.getElementById(id);
        }

        public void GoBack()
        {
            ie.GoBack();
        }

        public void OpenInternetExplorer(string url)
        {
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
    }
}

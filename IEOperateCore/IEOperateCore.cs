using IEOperateCore.IEBrowser;
using IEOperateCore.Interface;
using SHDocVw;
using mshtml;
using System.Threading;

namespace IEOperateCore
{
    public class IEOperateCore : IOperateDom
    {
        InternetExplorer ie;
        HTMLDocumentClass dom;
        public IEOperateCore() {
            ie = InternetExplorerFactory.GetInternetExplorer();
        }

        public void CloseInternetExplorer()
        {
            
            if (dom != null) {
                dom.close();
            }

            ie.Quit();
        }

        public HTMLDocumentClass GetDom()
        {

            while (ie.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                Thread.Sleep(500);
            }

            dom= (HTMLDocumentClass)ie.Document;
            return dom;
        }

      

        public HTMLInputElementClass GetInputElementByID(HTMLDocumentClass dom, string id)
        {
            return (HTMLInputElementClass)dom.getElementById(id);
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

using mshtml;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace IEOperateCore.Interface
{
    public interface IOperateDom: IDisposable
    {
        void SetInternetExplorerWindowPosition(int top, int left);
        void SetInternetExplorerWindowSize(int height, int width);
        void OpenInternetExplorer(string url);
        void GoBack();
        void Refresh();
        void CloseInternetExplorer();
        HTMLDocumentClass GetDom();
        T GetInputElementByID<T>(string id);
        IList<T> getElementByTagName<T>(string tagName);
        void SenKey(string keyBoard);
        IList<T> getElementByName<T>(string name);
        void SetIETabActivate( );
        
        IntPtr FindWindow(string lpClassName, string lpWindowName);
      
        List<HtmlElement> GetHtmlElementByTagName(IHTMLElement pieceDom, string tagName);
        bool InternetExplorerWindowIsReady(string url);
        /// <summary>
        /// 此方法是查找iframe下的a标签,查找规则是根据提供的searchKey查找指定的属性内容，默认直返回第一个被查找到的元素
        /// </summary>
        /// <param name="iframe"></param>
        /// <param name="searchKey"></param>
        /// <param name="isFindByTitle"></param>
        /// <param name="isFindByClassName"></param>
        /// <returns></returns>
        HTMLAnchorElementClass Get_Alink_From_IFrame( HTMLFrameElementClass iframe, string searchKey, bool isFindByTitle = true, bool isFindByClassName = false);
         
    }

   
}

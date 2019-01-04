using IEOperateLib.Common;
using mshtml;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace IEOperateLib.Interface
{
    public interface IOperateDom : IDisposable
    {
        bool IsMatchIEPage(string url);
        void SetInternetExplorerWindowPosition(int top, int left);
        void SetInternetExplorerWindowSize(int height, int width);
        void OpenInternetExplorer(string url);
        void GoBack();
        void Refresh();
        void CloseInternetExplorer();
        HTMLDocumentClass GetDom();
        T GetInputElementByID<T>(string id);
        IList<T> GetElementByTagName<T>(string tagName);
        void SenKey(string keyBoard);
        IList<T> GetElementByName<T>(string name);
        void SetIETabActivate();

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
        HTMLAnchorElementClass Get_Alink_From_IFrame(HTMLFrameElementClass iframe, string searchKey, bool isFindByTitle = true, bool isFindByClassName = false);

        HTMLTableClass FindTableFromDocument(IHTMLDocument2 doc, string tableId);

        HTMLTableClass FindTableFromDocument(HTMLDocument doc, string tableId);

        HtmlTable GetDataFromHtmlTable(HTMLTableClass tableCalss);

        void FrameNotificationBar_DownLoadFile_Save(string savePath = null, string windowTitle = null);

        void FrameNotificationBar_DownLoadFile_SaveAs(string savePath = null, string windowTitle = null);
    }
}

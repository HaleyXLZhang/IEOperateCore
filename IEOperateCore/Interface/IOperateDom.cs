using mshtml;
using SHDocVw;
using System;
using System.Collections.Generic;

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
    }

   
}

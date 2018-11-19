using mshtml;
using SHDocVw;
using System;

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
        T GetInputElementByID<T>(HTMLDocumentClass dom,string id);
        void SetIETabActivate( );
    }

   
}

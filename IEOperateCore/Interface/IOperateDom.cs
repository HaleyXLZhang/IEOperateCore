using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IEOperateCore.Interface
{
    public interface IOperateDom
    {
        void SetInternetExplorerWindowPosition(int top, int left);
        void SetInternetExplorerWindowSize(int height, int width);
        void OpenInternetExplorer(string url);
        void GoBack();
        void Refresh();
        void CloseInternetExplorer();
        HTMLDocumentClass GetDom();
        HTMLInputElementClass GetInputElementByID(HTMLDocumentClass dom,string id);

    }
}

using SHDocVw;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IEOperateCore.IEBrowser
{
    public class InternetExplorerFactory
    {
        private InternetExplorerFactory() { }
        private static InternetExplorer IE = null;
        public static  InternetExplorer GetInternetExplorer()
        {
            if (IE == null)
            {
                IE = new InternetExplorer();
                IE.Visible = true;
            }
            return IE;
        }
    }
}

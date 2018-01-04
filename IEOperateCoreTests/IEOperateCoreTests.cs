using Microsoft.VisualStudio.TestTools.UnitTesting;
using mshtml;

namespace IEOperateCore.Tests
{
    [TestClass()]
    public class IEOperateCoreTests
    {
        [TestMethod()]
        public void RefreshTest()
        {
            IEOperateCore ieCore = new IEOperateCore();

            ieCore.OpenInternetExplorer("www.baidu.com");
           // ieCore.Refresh();
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID(ieCore.GetDom(), "kw");
            searchTextBox.value = "China";
            HTMLInputElementClass searchButton = ieCore.GetInputElementByID(ieCore.GetDom(), "su");
            searchButton.click();
            ieCore.CloseInternetExplorer();
        }
    }
}
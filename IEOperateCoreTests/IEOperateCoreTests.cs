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
            //ieCore.Refresh();
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>(ieCore.GetDom(), "kw");
            searchTextBox.value = "China";
            HTMLInputElementClass searchButton = ieCore.GetInputElementByID<HTMLInputElementClass>(ieCore.GetDom(), "su");
            searchButton.click();
            
        }
        [TestMethod()]
        public void RefreshTestNew()
        {
            IEOperateCore ieCore = new IEOperateCore("www.baidu.com");     
            //ieCore.Refresh();
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>(ieCore.GetDom(), "kw");
            searchTextBox.value = "China";
            HTMLInputElementClass searchButton = ieCore.GetInputElementByID<HTMLInputElementClass>(ieCore.GetDom(), "su");
            searchButton.click();
           
        }
        [TestMethod()]
        public void Open360WebSite()
        {
            IEOperateCore ieCore = new IEOperateCore();
            ieCore.OpenInternetExplorer("https://hao.360.cn");
            //ieCore.Refresh();
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>(ieCore.GetDom(), "search-kw");
            searchTextBox.value = "China";
            HTMLButtonElementClass searchButton = ieCore.GetInputElementByID<HTMLButtonElementClass>(ieCore.GetDom(), "search-btn");
            searchButton.click();
            ieCore.CloseInternetExplorer();
        }

        [TestMethod()]
        public void SetTab360WebSite()
        {
            IEOperateCore ieCore = new IEOperateCore("https://hao.360.cn/");

            ieCore.SetIETabActivate();
        }
    }
}
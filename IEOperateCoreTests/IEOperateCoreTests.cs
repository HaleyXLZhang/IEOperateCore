using IEOperateCore.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using mshtml;
using System.Threading;
using System.Windows.Forms;

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
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>("kw");
            searchTextBox.value = "China";
            HTMLInputElementClass searchButton = ieCore.GetInputElementByID<HTMLInputElementClass>("su");
            searchButton.click();

        }
        [TestMethod()]
        public void RefreshTestNew()
        {
            IEOperateCore ieCore = new IEOperateCore("www.baidu.com");
            //ieCore.Refresh();
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>("kw");
            searchTextBox.value = "China";
            HTMLInputElementClass searchButton = ieCore.GetInputElementByID<HTMLInputElementClass>("su");
            searchButton.click();

        }
        [TestMethod()]
        public void Open360WebSite()
        {
            IEOperateCore ieCore = new IEOperateCore();
            ieCore.OpenInternetExplorer("https://hao.360.cn");
           
            HTMLInputElementClass searchTextBox = ieCore.GetInputElementByID<HTMLInputElementClass>("search-kw");

            
            searchTextBox.setActive();
           
            searchTextBox.value = "China";
            //HTMLButtonElementClass searchButton = ieCore.GetInputElementByID<HTMLButtonElementClass>("search-btn");
            searchTextBox.select();
            ieCore.SenKey(KeyBoard.Backspace);
            // searchButton.click();

            ieCore.SenKey(KeyBoard.Enter);

            // searchButton.FireEvent("submit");

            // ieCore.CloseInternetExplorer();
        }

        [TestMethod()]
        public void SetTab360WebSite()
        {
            IEOperateCore ieCore = new IEOperateCore("https://hao.360.cn/");

            ieCore.SetIETabActivate();
        }
        [TestMethod()]
        public void GetelementsByTagNmae()
        {
            IEOperateCore ieCore = new IEOperateCore("https://hao.360.cn/");

          var list=  ieCore.getElementByTagName<HTMLLIElementClass>("li");
        }
    }
}
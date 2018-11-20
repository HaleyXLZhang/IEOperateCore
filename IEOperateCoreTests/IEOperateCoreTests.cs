using IEOperateCore.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using mshtml;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using System;
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


            //searchTextBox.setActive();
            searchTextBox.value = "China";
            //searchTextBox.select();
            //ieCore.SenKey(KeyBoard.Backspace);
            //ieCore.SenKey(KeyBoard.Enter);


            HTMLButtonElementClass searchButton = ieCore.GetInputElementByID<HTMLButtonElementClass>("search-btn");
            searchButton.click();
            //searchButton.FireEvent("click");
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

            var list = ieCore.getElementByTagName<HTMLLIElementClass>("li");

        }
        [TestMethod()]
        public void MatchWeb()
        {
            IEOperateCore ieCore = new IEOperateCore("http://psgpms.chinasofti.com/main.jsp");

            HTMLFrameElementClass iframe = ieCore.GetInputElementByID<HTMLFrameElementClass>("ManagerTopFrame");

            IHTMLDocument2 doc = iframe.contentWindow.document;

            HTMLBodyClass bodyClass = (HTMLBodyClass)doc.body;

            IHTMLElementCollection aLinkElementCollection = bodyClass.getElementsByTagName("a");

            foreach (HTMLAnchorElementClass item in aLinkElementCollection)
            {
                if (item.title.Equals("退出系统"))
                {


                    item.click();

                    break;
                }
            }

            IntPtr ieDialogWindowHWND = ieCore.FindWindow(null, "来自网页的消息");
            Thread.Sleep(2000);
            IntPtr childHwnd = Win32.FindWindowEx(ieDialogWindowHWND, IntPtr.Zero, null, "取消");   //获得查询按钮的句柄   
            if (childHwnd != IntPtr.Zero)
            {
                Win32.SendMessage(childHwnd, Win32.BM_CLICK, IntPtr.Zero, IntPtr.Zero);//发送点击按钮的消息                                             
            }
        }

        [TestMethod()]
        public void TestOperateCaculate()
        {
            //FindWindow 参数一是进程名 参数2是 标题名
            IntPtr calculatorHandle = Win32.FindWindow(null, "计算器");
            //判断是否找到
            if (calculatorHandle == IntPtr.Zero)
            {
                //  MessageBox.Show("没有找到!");
                return;
            }
            // 然后使用SetForegroundWindow函数将这个窗口调到最前。
            Win32.SetForegroundWindow(calculatorHandle);
            //发送按键



            SendKeys.SendWait("2");

            SendKeys.SendWait("*");

            SendKeys.SendWait("11");

            SendKeys.SendWait("=");
        }

    }
}
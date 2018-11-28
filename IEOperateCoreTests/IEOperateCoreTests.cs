using IEOperateCore.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using mshtml;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Windows.Forms;
using System.Text;
using Microsoft.VisualStudio.CommandBars;
using System.Windows.Automation;
using Microsoft.Win32;
using System.Text.RegularExpressions;

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
        public void GeAlinkFromFrame()
        {

            IEOperateCore ieCore = new IEOperateCore("http://psgis.chinasofti.com/oa/SignOnServlet");
            HTMLInputElementClass userName = ieCore.getElementByTagName<HTMLInputElementClass>("input").ToList().Find(item => item.name.Equals("userName"));
            HTMLInputElementClass passWord = ieCore.getElementByTagName<HTMLInputElementClass>("input").ToList().Find(item => item.name.Equals("password"));
            HTMLImgClass imgLogin = ieCore.getElementByTagName<HTMLImgClass>("img").ToList().Find(item => item.name.Equals("login_r3_c3"));
            userName.value = string.Empty;
            passWord.value = string.Empty;
            userName.value = "E100086376";
            passWord.value = "Qa20080607";
            imgLogin.click();
            Thread.Sleep(3500);
            ieCore = new IEOperateCore("http://psgis.chinasofti.com/oa/portal");
            HTMLDivElementClass btnPMS = ieCore.getElementByTagName<HTMLDivElementClass>("div").ToList().Find(item => item.className.Equals("channel app_PMS系统"));
            Thread.Sleep(4000);
            btnPMS.click();
            Thread.Sleep(5000);

            ieCore = new IEOperateCore("http://psgpms.chinasofti.com/main.jsp");


            HTMLFrameElementClass midIframe = ieCore.GetInputElementByID<HTMLFrameElementClass>("ManagerLeftFrame");




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


        //

        [TestMethod()]
        public void CaptureTableData()
        {

            StringBuilder buffer = new StringBuilder();

            IEOperateCore ieCore = new IEOperateCore("http://psgpms.chinasofti.com/main.jsp");

            HTMLFrameElementClass iframe = ieCore.GetInputElementByID<HTMLFrameElementClass>("mainFrame");

            IHTMLDocument2 doc = iframe.contentWindow.document;

            HTMLBodyClass bodyClass = (HTMLBodyClass)doc.body;

            IHTMLElementCollection tableElementCollection = bodyClass.getElementsByTagName("table");

            foreach (HTMLTableClass table in tableElementCollection)
            {
                if (null != table.id && table.id.Equals("rossTableId_TASK"))
                {
                    foreach (HTMLTableRowClass row in table.rows)
                    {
                        foreach (HTMLTableCellClass cell in row.cells)
                        {
                            if (null != cell.innerText)
                                buffer.Append(cell.innerText + "        ");
                        }
                        buffer.Append("/r/n");

                    }
                    break;
                }
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




        [TestMethod()]
        public void OperateDirectUI()
        {
            IEOperateCore ieCore = new IEOperateCore("https://dl.pconline.com.cn/download/63040-1.html");

            Thread.Sleep(2000);

            IList<HTMLAnchorElementClass> aLinks = ieCore.getElementByTagName<HTMLAnchorElementClass>("a");

            foreach (var a in aLinks)
            {
                if (null != a.className && a.className.Equals("btn btn-gray"))
                {
                    a.click();
                    Thread.Sleep(3000);
                    break;
                }
            }


            ieCore.FrameNotificationBar_DownLoadFile_Save();

            ieCore.CloseInternetExplorer();
        }

        [TestMethod()]
        public void OperateDownloadFrameNotificationBar()
        {
            IEOperateCore ieCore = new IEOperateCore("https://dl.pconline.com.cn/download/63040-1.html");

            Thread.Sleep(2000);

           

            // string original = "要运行或保存来自 ftp-idc.pconline.com.cn 的 70.0.3538.102_chrome_installer_32.exe (49.8 MB) 吗?";
            string original = "Do you want to run or save 70.0.3538.102_chrome_installer_32.exe (49.8 MB) from ftp-idc.pconline.com.cn?";


            string patternStr = "(?<= 的 | save )+.+.[A-Za-z]+ \\(\\d+.*\\d* [A-Za-z]+\\)";

            Regex regex = new Regex(patternStr);

            Match title = regex.Match(original);

            string fileSizePattern = "\\(\\d+.*\\d* [A-Za-z]+\\)";

            regex = new Regex(fileSizePattern);

            Match fileSizeMatch = regex.Match(title.Value);

            string fileSize = fileSizeMatch.Value;

            string patternDownloadFileName = "(?<= 的 | save )+.+.[A-Za-z]+ (?=\\(\\d+.*\\d* [A-Za-z]+\\))";

            regex = new Regex(patternDownloadFileName);

            string fileName = regex.Match(original).Value;

            IntPtr parentHandle = Win32.FindWindow("IEFrame", null);

            AutomationElementCollection parentElements = AutomationElement.FromHandle(parentHandle).FindAll(TreeScope.Children, Condition.TrueCondition);

            foreach (AutomationElement parentElement in parentElements)
            {
                // Identidfy Download Manager Window in Internet Explorer
                if (parentElement.Current.ClassName == "Frame Notification Bar")
                {
                    AutomationElementCollection childElements = parentElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    // Idenfify child window with the name Notification Bar or class name as DirectUIHWND 
                    foreach (AutomationElement childElement in childElements)
                    {
                        if (childElement.Current.Name == "Notification bar" || childElement.Current.ClassName == "DirectUIHWND")
                        {
                            AutomationElementCollection downloadCtrls = childElement.FindAll(TreeScope.Descendants, Condition.TrueCondition);

                            AutomationElement fileNameText = childElement.FindFirst(TreeScope.Descendants,
                                                                        new AndCondition(
                                                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                                                                        new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                                                        ));
                            dynamic value = (ValuePattern)fileNameText.GetCurrentPattern(ValuePattern.Pattern);

                            string content = (string)value.Current.Value;
                            //要运行或保存来自 ftp-idc.pconline.com.cn 的 70.0.3538.102_chrome_installer_32.exe (49.8 MB) 吗?
                            //Do you want to run or save 70.0.3538.102_chrome_installer_32.exe (49.8 MB) from ftp-idc.pconline.com.cn?

                            foreach (AutomationElement ctrlButton in downloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                //另存为
                                if (ctrlButton.Current.Name.Equals("关闭") || ctrlButton.Current.Name.ToLower().Equals("Close"))
                                {
                                    AutomationElementCollection contentTextChilds = ctrlButton.FindAll(TreeScope.Children, Condition.TrueCondition);
                                    foreach (AutomationElement text in contentTextChilds)
                                    {
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
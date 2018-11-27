using IEOperateCore.IEBrowser;
using IEOperateCore.Interface;
using SHDocVw;
using mshtml;
using System.Threading;
using System;
using IEOperateCore.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Automation;
using Microsoft.Win32;
using System.IO;

namespace IEOperateCore
{
    public class IEOperateCore : IOperateDom
    {
        public IntPtr HWND = new IntPtr(0);
        public InternetExplorer IE;
        private HTMLDocumentClass dom;

        public IEOperateCore()
        {
            IE = InternetExplorerFactory.GetInternetExplorer();
            HWND = new IntPtr(IE.HWND);

        }
        public IEOperateCore(string url)
        {
            IE = InternetExplorerFactory.GetInternetExplorer(url);
            HWND = new IntPtr(IE.HWND);
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                try
                {
                    dom = (HTMLDocumentClass)IE.Document;
                }
                catch (Exception e)
                {
                    Thread.Sleep(1000);
                    continue;
                }
                if (dom.readyState.Equals("complete")) break;

                if (loopCount > 2000)
                    throw new Exception("open " + url + " timeout!");
                Thread.Sleep(500);
                loopCount++;
            }
        }
        public void SetIETabActivate()
        {
            new TabActivator((IntPtr)IE.HWND).ActivateByTabsUrl(IE.LocationURL);
        }
        public void CloseInternetExplorer()
        {

            if (dom != null)
                dom.close();
            if (IE != null)
                IE.Quit();
        }

        public void Dispose()
        {
            CloseInternetExplorer();
        }
        public bool InternetExplorerWindowIsReady(string url)
        {
            SHDocVw.ShellWindows ieTabs = new SHDocVw.ShellWindows();

            foreach (SHDocVw.InternetExplorer ieTab in ieTabs)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ieTab.FullName).ToLower();

                if (filename.Equals("iexplore") && ieTab.LocationURL.Equals(url))
                {
                    return !ieTab.Busy;
                }
            }
            return false;
        }

        public HTMLDocumentClass GetDom()
        {
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                if (loopCount > 2000)
                    throw new Exception("Get dom timeout!");
                Thread.Sleep(500);
                loopCount++;
            }

            dom = (HTMLDocumentClass)IE.Document;


            return dom;
        }

        public T GetInputElementByID<T>(string id)
        {


            GetDom();
            return (T)dom.getElementById(id);
        }

        public void SenKey(string keyBoard)
        {
            SendKeys.SendWait(keyBoard);
        }

        public void GoBack()
        {
            IE.GoBack();
        }

        public void OpenInternetExplorer(string url)
        {
            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_TOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            Win32.SetWindowPos(new IntPtr(IE.HWND), (IntPtr)Win32.hWndInsertAfter.HWND_NOTTOPMOST, 0, 0, 0, 0, Win32.TOPMOST_FLAGS);

            IE.Navigate(url);
            int loopCount = 0;
            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                try
                {
                    dom = (HTMLDocumentClass)IE.Document;
                }
                catch (Exception e)
                {
                    Thread.Sleep(1000);
                    continue;
                }
                if (dom.readyState.Equals("complete")) break;
                if (loopCount > 2000)
                    throw new Exception("Get " + url + " timeout!");
                Thread.Sleep(500);
                loopCount++;
            }
        }

        public void Refresh()
        {

            while (IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
                Thread.Sleep(500);
            }
            IE.Refresh();
        }

        public void SetInternetExplorerWindowPosition(int top, int left)
        {
            IE.Top = top;
            IE.Left = left;
        }

        public void SetInternetExplorerWindowSize(int height, int width)
        {
            IE.Height = height;
            IE.Width = width;
        }

        public IList<T> getElementByTagName<T>(string tagName)
        {
            IList<T> getCollection = new List<T>();

            GetDom();

            IHTMLElementCollection collection = dom.getElementsByTagName(tagName);

            foreach (IHTMLElement elem in collection)
            {
                if (null != elem)
                    getCollection.Add((T)elem);
            }
            return getCollection;
        }

        public IList<T> getElementByName<T>(string name)
        {
            IList<T> getCollection = new List<T>();

            GetDom();

            IHTMLElementCollection collection = dom.getElementsByName(name);

            foreach (IHTMLElement elem in collection)
            {
                if (null != elem)
                    getCollection.Add((T)elem);
            }
            return getCollection;
        }


        public IntPtr FindWindow(string lpClassName, string lpWindowName)
        {
            return Win32.FindWindow(lpClassName, lpWindowName);
        }
        public List<HtmlElement> GetHtmlElementByTagName(IHTMLElement pieceDom, string tagName)
        {

            List<HtmlElement> getCollection = new List<HtmlElement>();

            //HtmlElementCollection findCollection = pieceDom.innerText(tagName);


            //foreach (HtmlElement element in findCollection)
            //{
            //    if (null != element)
            //        getCollection.Add(element);
            //}
            return getCollection;
        }


        /// <summary>
        /// 此方法是查找iframe下的a标签,查找规则是根据提供的searchKey查找指定的属性内容，默认直返回第一个被查找到的元素
        /// </summary>
        /// <param name="iframe"></param>
        /// <param name="searchKey"></param>
        /// <param name="isFindByTitle"></param>
        /// <param name="isFindByClassName"></param>
        /// <returns></returns>
        public HTMLAnchorElementClass Get_Alink_From_IFrame(HTMLFrameElementClass iframe, string searchKey, bool isFindByTitle = true, bool isFindByClassName = false)
        {
            HTMLAnchorElementClass aLink = null;
            IHTMLElementCollection aLinkElementCollection;
            IHTMLDocument2 doc;
            HTMLBodyClass bodyClass;
            doc = iframe.contentWindow.document;
            bodyClass = (HTMLBodyClass)doc.body;
            aLinkElementCollection = bodyClass.getElementsByTagName("a");
            foreach (HTMLAnchorElementClass item in aLinkElementCollection)
            {
                if (isFindByTitle)
                {
                    if (null != item.title && item.title.Equals(searchKey))
                    {
                        aLink = item;
                        break;
                    }
                }
                else if (isFindByClassName)
                {
                    if (null != item.className && item.className.Equals(searchKey))
                    {
                        aLink = item;
                        break;
                    }
                }
            }
            return aLink;
        }
        public HTMLTableClass FindTableFromDocument(IHTMLDocument2 doc, string tableId)
        {

            HTMLTableClass findTable = null;

            HTMLBodyClass bodyClass = (HTMLBodyClass)doc.body;

            IHTMLElementCollection tableElementCollection = bodyClass.getElementsByTagName("table");

            foreach (HTMLTableClass table in tableElementCollection)
            {
                if (null != table.id && table.id.Equals(tableId))
                {
                    findTable = table;
                    break;
                }
            }
            return findTable;
        }

        public HTMLTableClass FindTableFromDocument(HTMLDocument doc, string tableId)
        {
            HTMLTableClass findTable = null;

            HTMLBodyClass bodyClass = (HTMLBodyClass)doc.body;

            IHTMLElementCollection tableElementCollection = bodyClass.getElementsByTagName("table");

            foreach (HTMLTableClass table in tableElementCollection)
            {
                if (null != table.id && table.id.Equals(tableId))
                {
                    findTable = table;
                    break;
                }
            }
            return findTable;
        }

        public HtmlTable GetDataFromHtmlTable(HTMLTableClass tableCalss)
        {
            HtmlTable table = new HtmlTable();
            HtmlTableRow headerRow = new HtmlTableRow();
            int rowIndex = 0;
            foreach (HTMLTableRowClass row in tableCalss.rows)
            {
                // table header
                if (rowIndex == 0)
                {
                    int columnIndex = 0;
                    foreach (HTMLTableCellClass cell in row.cells)
                    {
                        if (null != cell.innerText)
                        {
                            HtmlTableCell newCell = new HtmlTableCell();
                            newCell.CellName = cell.innerText;
                            newCell.CellColumnIndex = columnIndex;
                            newCell.CellValue = null;
                            headerRow.Cells.Add(newCell);
                        }
                        columnIndex++;
                    }
                }
                else
                {
                    // table body
                    HtmlTableRow newRow = new HtmlTableRow();
                    newRow.RowIndex = --rowIndex;
                    int columnIndex = 0;
                    foreach (HTMLTableCellClass cell in row.cells)
                    {
                        if (null != cell.innerText)
                        {
                            HtmlTableCell newCell = new HtmlTableCell();
                            newCell.CellName = headerRow.Cells[columnIndex].CellName;
                            newCell.CellColumnIndex = columnIndex;
                            newCell.CellValue = cell.innerText;
                            newRow.Cells.Add(newCell);
                        }
                        columnIndex++;
                    }
                    table.HtmlTableRows.Add(newRow);
                }
                rowIndex++;
            }
            return table;
        }

        public void FrameNotificationBar_DownLoadFile_Save(string savePath = null, string windowTitle = null)
        {
            SaveFileFrameNotificationBar(savePath,windowTitle);
        }

        public void FrameNotificationBar_DownLoadFile_SaveAs(string savePath = null, string windowTitle = null)
        {
            SaveAsFileFrameNotificationBar(savePath, windowTitle);
        }

        private void SaveAsFileFrameNotificationBar(string saveFile = null, string windowTitle = null)
        {
            String path = String.Empty;
            IntPtr parentHandle = Win32.FindWindow("IEFrame", windowTitle);
            var parentElements = AutomationElement.FromHandle(parentHandle).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement parentElement in parentElements)
            {
                // Identidfy Download Manager Window in Internet Explorer
                if (parentElement.Current.ClassName == "Frame Notification Bar")
                {
                    var childElements = parentElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    // Idenfify child window with the name Notification Bar or class name as DirectUIHWND 
                    foreach (AutomationElement childElement in childElements)
                    {
                        if (childElement.Current.Name == "Notification bar" || childElement.Current.ClassName == "DirectUIHWND")
                        {
                            var downloadCtrls = childElement.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                            foreach (AutomationElement ctrlButton in downloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                //另存为
                                if (ctrlButton.Current.Name.ToLower() == "")
                                {
                                    var saveSubMenu = ctrlButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                    saveSubMenu.Invoke();
                                    SendKeys.SendWait("{DOWN}");
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{ENTER}");

                                    Thread.Sleep(2000);

                                    IntPtr saveMenuHandle = Win32.FindWindow(null, "另存为");

                                    if (saveMenuHandle == IntPtr.Zero)

                                        saveMenuHandle = Win32.FindWindow(null, "Save As");

                                    var subMenuItems = AutomationElement.FromHandle(saveMenuHandle).FindAll(TreeScope.Children, Condition.TrueCondition);
                                    foreach (AutomationElement item in subMenuItems)
                                    {
                                        if (item.Current.ClassName.Equals("DUIViewWndClassName"))
                                        {
                                            AutomationElement fileNameText = item.FindFirst(TreeScope.Descendants,
                                                                                        new AndCondition(
                                                                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                                                                                        new PropertyCondition(AutomationElement.ClassNameProperty, "Edit")
                                                                                        ));
                                            if (null != saveFile)
                                                Win32.SendMessage(new IntPtr(fileNameText.Current.NativeWindowHandle), Win32.WM_SETTEXT, IntPtr.Zero, saveFile);
                                        }
                                        if (item.Current.Name.Equals("保存(S)") || item.Current.Name.ToLower().Equals("save(s)"))
                                        {
                                            var saveAsMenuItem = item.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                            saveAsMenuItem.Invoke();
                                            break;
                                        }
                                    }
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
            }

            if (null == saveFile)
            {
                RegistryKey rKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main");
                if (rKey != null)
                    path = (String)rKey.GetValue("Default Download Directory");
                if (String.IsNullOrEmpty(path))
                    path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads";
            }
            else
            {
                path = saveFile;
            }

            int waitCount = 1800;
            do
            {
                Thread.Sleep(1000);
                waitCount--;
                if (waitCount < 0) throw new Exception(string.Format("Download Failed: {0}", path));
            } while (!File.Exists(path));

            CloseFrameNotificationBar(windowTitle);
        }
        private void SaveFileFrameNotificationBar(string saveFile = null, string windowTitle = null)
        {
            String path = String.Empty;
            IntPtr parentHandle = Win32.FindWindow("IEFrame", windowTitle);
            var parentElements = AutomationElement.FromHandle(parentHandle).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement parentElement in parentElements)
            {
                // Identidfy Download Manager Window in Internet Explorer
                if (parentElement.Current.ClassName == "Frame Notification Bar")
                {
                    var childElements = parentElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    // Idenfify child window with the name Notification Bar or class name as DirectUIHWND 
                    foreach (AutomationElement childElement in childElements)
                    {
                        if (childElement.Current.Name == "Notification bar" || childElement.Current.ClassName == "DirectUIHWND")
                        {
                            var downloadCtrls = childElement.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                            foreach (AutomationElement ctrlButton in downloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                //另存为
                                if (ctrlButton.Current.Name.Equals("保存") || ctrlButton.Current.Name.ToLower().Equals("save"))
                                {
                                    var saveSubMenu = ctrlButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                    saveSubMenu.Invoke();
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
            }

            if (null == saveFile)
            {
                RegistryKey rKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main");
                if (rKey != null)
                    path = (String)rKey.GetValue("Default Download Directory");
                if (String.IsNullOrEmpty(path))
                    path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads";
            }
            else
            {
                path = saveFile;
            }

            int waitCount = 1800;
            do
            {
                Thread.Sleep(1000);
                waitCount--;
                if (waitCount < 0) throw new Exception(string.Format("Download Failed: {0}", path));
            } while (!File.Exists(path));

            CloseFrameNotificationBar(windowTitle);
        }
        private void CloseFrameNotificationBar(string windowTitle = null)
        {
            IntPtr parentHandle = Win32.FindWindow("IEFrame", windowTitle);

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

                            foreach (AutomationElement ctrlButton in downloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                //另存为
                                if (ctrlButton.Current.Name.Equals("关闭") || ctrlButton.Current.Name.ToLower().Equals("Close"))
                                {
                                    var closeButton = ctrlButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                                    closeButton.Invoke();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.IO;
using mshtml;

namespace AutoBroswer
{
    class SimulateTab : Form
    {
        System.Windows.Forms.Timer timeDown = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer timeUp = new System.Windows.Forms.Timer();

        DateTime jobExpireTime; //任务超时时间 15M
        DateTime pageExpireTime;//摸个页面超时时间
        DateTime openURLExpireTime;//加载URL超时

        System.Windows.Forms.Timer moniterTimer = new System.Windows.Forms.Timer();
        //System.Windows.Forms.Timer pageMoniterTimer = new System.Windows.Forms.Timer();
        //System.Windows.Forms.Timer expireTimer = new System.Windows.Forms.Timer();

        //System.Threading.Timer m_openURLTimer;
        public int randMoveInterval = 50;

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool BlockInput(bool block);
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        
        [DllImport("user32.dll", CharSet=CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName,int nMaxCount);

        //string keyWord;
        AutoBroswerForm.STKeyInfo keyInfo;
        AutoBroswerForm autoBroswerFrom;
        enum ECurrentStep
        {
            ECurrentStep_Load,
            ECurrentStep_Search,
            ECurrentStep_Visit_Compare,
            ECurrentStep_Visit_Me_Main,//访问主宝贝
            ECurrentStep_Visit_Me_MainPage,//访问主页
            ECurrentStep_Visit_Me_Other//访问其他宝贝
        }

        ECurrentStep m_currentStep = ECurrentStep.ECurrentStep_Load;

        int currentScrolBarPos = 0;
        private string m_uaString;
        public int m_iMainItemStopMin;//主宝贝停留时间
        public int m_iMainItemStopMax;//
        public int m_iOhterItemStopMin;//次宝贝停留时间
        public int m_iOtherItemStopMax;

        public int m_iOtherItemStopTime;//其它家宝贝时间，随机在20-30s
        const int millSeconds = 1000;
        const int OPENURLTIMEOUT = 30 * 1000;
        const int ImpossibleTime = 15 * 60 * 1000;//不可能的超时时间
        private bool isOpenningURL = false;
        ExtendedWebBrowser InitialTabBrowser;

        private HtmlElement m_myItemElement;//在搜索页的ELEMENT,
        private HtmlElement m_myMainPageElement;//在在主宝贝页面中的首页
        //private List<HtmlElement> m_randItemElement;//其它随机宝贝,<p class="pic-box">
        private string[] m_clickLinkItem = { "评价详情", "成交记录", "宝贝详情" };
        private string[] m_clickMainPageItem = { "按销量", "查看所有宝贝", "进入店铺", "按收藏" };
        private string[] m_clickSpanItem = { "物流运费", "销　　量", "评　　价", "宝贝类型", "支　　付" };

        private List<HtmlElement> m_mainItemClickElement;//详情页三个模块点击
        private List<HtmlElement> m_mainItemSpanElement;//详情页Span模块点击

        private string m_otherItemPattern = @"http://item.taobao.com/item.htm?(.*)id=(\d{11})$";
        private string m_otherTianMallPattern = @"http://detail.tmall.com/item.htm?(.*)id=(\d{11})$";

        public bool isNormalQuit = false;
        private int m_randCompCount = 0;//随机货比三家个数
        private int m_randDeepItemCount = 0;//访问深度
        Regex otherItemRegex;
        Regex otherTianMallRegex;

        public bool initValue()
        {
            otherItemRegex = new Regex(m_otherItemPattern);
            otherTianMallRegex = new Regex(m_otherTianMallPattern);
            m_iMainItemStopMin = autoBroswerFrom.getMainItemMinTime();
            m_iMainItemStopMax = autoBroswerFrom.getMainItemMaxTime();
            m_iOhterItemStopMin = autoBroswerFrom.getOtherItemMinTime();
            m_iOtherItemStopMax = autoBroswerFrom.getOtherItemMaxTime();

            m_mainItemClickElement = new List<HtmlElement>();
            m_mainItemSpanElement = new List<HtmlElement>();

            //pageMoniterTimer.Tick += new EventHandler(PageMoniterTimeEvent);

            if (autoBroswerFrom.isCompareRandCB())
            {
                m_randCompCount = autoBroswerFrom.rndGenerator.Next(1, 3);
            }

            m_randDeepItemCount = autoBroswerFrom.getVisitDeep();
            if (autoBroswerFrom.isVisitDeepRand())
            {
                m_randDeepItemCount = autoBroswerFrom.rndGenerator.Next(1, 5);
            }

            return true;
        }
        public SimulateTab(AutoBroswerForm.STKeyInfo _keyInfo, string uaString, AutoBroswerForm _AutoBroswer, int expireTime)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopLevel = false;
            this.Size = _AutoBroswer.webBrowserPanel.Size;
            m_uaString = uaString;

            if (_AutoBroswer.webBrowserPanel.InvokeRequired)
            {
                _AutoBroswer.webBrowserPanel.BeginInvoke(new MethodInvoker(delegate
                {
                    this.Parent = _AutoBroswer.webBrowserPanel;
                    
                }));
            }
            InitialTabBrowser = new ExtendedWebBrowser()
            {
                Parent = this,
                Dock = DockStyle.Fill,
                ScriptErrorsSuppressed = true,
                UserAgent = uaString
            };

            timeDown.Interval = 100;
            timeUp.Interval = 100;
            InitialTabBrowser.Visible = true;
            InitialTabBrowser.NavigateError += new AutoBroswer.ExtendedWebBrowser.WebBrowserNavigateErrorEventHandler(InitialTabBrowser_NavigateError);
            if (_keyInfo.isZTCClick())
            {
                timeDown.Tick += new EventHandler(timeDownZTC_Tick);
                timeUp.Tick += new EventHandler(timeUpZTC_Tick);
                InitialTabBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(InitialTabBrowser_ZTCDocumentCompleted);
            }
            else
            {
                timeDown.Tick += new EventHandler(timeDown_Tick);
                timeUp.Tick += new EventHandler(timeUp_Tick);
                InitialTabBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(InitialTabBrowser_DocumentCompleted);
            }
            InitialTabBrowser.ProgressChanged += new WebBrowserProgressChangedEventHandler(InitialTabBrowser_ProgressChangedForSomething);

                        
            

            jobExpireTime = DateTime.Now.AddMilliseconds(expireTime);


            keyInfo = _keyInfo;
            autoBroswerFrom = _AutoBroswer;
            initValue();

            InitialTabBrowser.Navigate("http://www.taobao.com/");
            isOpenningURL = true;
            openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);

            moniterTimer.Tick += new EventHandler(TimerTick);
            moniterTimer.Interval = 1000;
            moniterTimer.Enabled = true;
            moniterTimer.Start();
            Application.DoEvents();
        }

        public bool searchBroswer(string keyword)
        {
            try
            {
                HtmlDocument document = ((HtmlDocument)InitialTabBrowser.Document);
                File.WriteAllText("search.txt", document.Body.OuterHtml.ToString(), Encoding.Default);
                HtmlElement textArea = document.GetElementById("q");
                if (textArea == null)
                {
                    FileLogger.Instance.LogInfo("找不到搜索输.入框标签");
                    return false;
                }
                FileLogger.Instance.LogInfo("before setInnerText");
                textArea.InnerText = keyword;
                FileLogger.Instance.LogInfo("after setInnerText");

                var elements = document.GetElementsByTagName("button");
                foreach (HtmlElement searchBTN in elements)
                {
                    string className = searchBTN.GetAttribute("className");
                    if (className == "btn-search" || searchBTN.InnerText == "搜 索")
                    {
                        FileLogger.Instance.LogInfo("found search button");
                        Rectangle rect = wbElementMouseSimulate.GetElementRect(document.Body.DomElement as mshtml.IHTMLElement, searchBTN.DomElement as mshtml.IHTMLElement);
                        Point p = new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2);
                        RandMove(InitialTabBrowser.Handle, 1000, rect);
                        ClickOnPointInClient(InitialTabBrowser.Handle, p);
                        isOpenningURL = true;
                        m_currentStep = ECurrentStep.ECurrentStep_Search;
                        return true;
                    }
                }
                FileLogger.Instance.LogInfo("找不到搜索按钮标签");
                return false;
            }
            catch (System.Exception ex)
            {
                FileLogger.Instance.LogInfo("CatchError:" + ex.Message);
            }
            return false;
        }
        private Point GetElementPosition(HtmlElement current_element)
        {
            int x_add = current_element.OffsetRectangle.Width;
            int y_add = current_element.OffsetRectangle.Height;
            int x = current_element.OffsetRectangle.Left;
            int y = current_element.OffsetRectangle.Top;
            while ((current_element = current_element.Parent) != null)
            {
                x += current_element.OffsetRectangle.Left;
                y += current_element.OffsetRectangle.Top;
            }

            y -= (InitialTabBrowser.Location.Y);

            return new Point(x + (x_add / 2), y + (y_add / 2));
        }
        private void SetCursorPos(int p, int p_2)
        {
            throw new NotImplementedException();
        }
        private void Window_Error(object sender,  HtmlElementErrorEventArgs e)
        {
            // Ignore the error and suppress the error dialog box. 
            e.Handled = true;
        }
        public void InitialTabBrowser_NavigateError(object sender, WebBrowserNavigateErrorEventArgs e)
        {
            // Display an error message to the user.
            //MessageBox.Show("Cannot navigate to " + e.Url);
            FileLogger.Instance.LogInfo("load fail:" + e.Url + " Status:" + e.StatusCode);
            isNormalQuit = true;
            this.ShutDownWinForms();
        }
        public void InitialTabBrowser_ProgressChangedForSomething(object sender, WebBrowserProgressChangedEventArgs e)
        {
            
            if (((WebBrowser)sender).ReadyState == WebBrowserReadyState.Complete)
            {
                //System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                //messageBoxCS.AppendFormat("{0} = {1}", "CurrentProgress", e.CurrentProgress);
                //messageBoxCS.AppendLine();
                //messageBoxCS.AppendFormat("{0} = {1}", "MaximumProgress", e.MaximumProgress);
                //messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ProgressChanged Event");
                
                //doMainJob();
            }
        }
        void InitialTabBrowser_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            FileLogger.Instance.LogInfo("Navigating URL:" + e.Url);
        }
        void InitialTabBrowser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            FileLogger.Instance.LogInfo("Navigated URL:" + e.Url);
        }
        public void InitialTabBrowser_ZTCDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            HtmlElement head = InitialTabBrowser.Document.GetElementsByTagName("head")[0];
            HtmlElement testScript = InitialTabBrowser.Document.CreateElement("script");
            IHTMLScriptElement element = (IHTMLScriptElement)testScript.DomElement;
            element.text = wbElementMouseSimulate.m_simulateJS;

            head.AppendChild(testScript);
            doMainJobZTC();
        }
        public void InitialTabBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            HtmlElement head = InitialTabBrowser.Document.GetElementsByTagName("head")[0];
            HtmlElement testScript = InitialTabBrowser.Document.CreateElement("script");
            IHTMLScriptElement element = (IHTMLScriptElement)testScript.DomElement;
            element.text = wbElementMouseSimulate.m_simulateJS;

            head.AppendChild(testScript); 
            doMainJob();
        }

        private string lastURL = "";
        public void doMainJob()
        {
            bool bRet = false;
            FileLogger.Instance.LogInfo("当前文档状态:" + this.InitialTabBrowser.ReadyState);
            if ((m_currentStep != ECurrentStep.ECurrentStep_Load && m_currentStep != ECurrentStep.ECurrentStep_Search) 
                && this.InitialTabBrowser.ReadyState != WebBrowserReadyState.Complete)
                return;
            string innerHtml = this.InitialTabBrowser.Document.Body.InnerHtml;
            string str1 = this.InitialTabBrowser.Document.Url.ToString();
            FileLogger.Instance.LogInfo("当前步骤:" + m_currentStep);
            //FileLogger.Instance.LogInfo("Cookie:" + InitialTabBrowser.Document.Cookie);
            autoBroswerFrom.URLTextBox.Text = str1;


            //将所有的链接的目标，指向本窗体   
            foreach (HtmlElement archor in this.InitialTabBrowser.Document.Links)
            {
                archor.SetAttribute("target", "_top");
            }  
            //将所有的FORM的提交目标，指向本窗体
            foreach (HtmlElement form in this.InitialTabBrowser.Document.Forms)
            {
                form.SetAttribute("target", "_top"); 
            }

            if (m_currentStep == ECurrentStep.ECurrentStep_Load)
            {
                DateTime dateExpire = DateTime.Parse("2013-12-25 12:30:01");
                if (DateTime.Now > dateExpire)
                {
                    MessageBox.Show("未知错误，可能淘宝又变标签了，请联系作者", "出错啦！");
                    return;
                }
                //bRet = searchBroswer(keyWord);
                if (str1.IndexOf("http://www.taobao.com/") > -1 && innerHtml.IndexOf("淘宝网首页") > -1)
                {
                    this.InitialTabBrowser.Document.GetElementById("q").InnerText = keyInfo.m_keyword;
                    this.InitialTabBrowser.Document.GetElementById("J_TSearchForm").InvokeMember("submit");
                    isOpenningURL = true;
                    m_currentStep = ECurrentStep.ECurrentStep_Search;
                    openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    lastURL = str1;
                }
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Search)
            {
                if (str1.IndexOf("http://s.taobao.com/search") > -1 && innerHtml.IndexOf("所有分类") > -1)
                {
                    SetTimerDownEnable(50);
                    //searchInPage();
                    autoBroswerFrom.currentStep.Text = "查找宝贝";
                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    isOpenningURL = false;
                    lastURL = str1;
                }
                
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Compare)
            {
                if (str1.IndexOf("http://s.taobao.com/search") > -1 && innerHtml.IndexOf("所有分类") > -1)
                {
                    SetTimerDownEnable(5);

                    autoBroswerFrom.currentStep.Text = "货比三家回搜索页";

                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    isOpenningURL = false;
                    //pageMoniterTimer.Enabled = false;
                    //pageMoniterTimer.Stop();
                }
                else if ((str1.IndexOf("http://item.taobao.com") > -1 || str1.IndexOf("http://detail.tmall.com") > -1) && innerHtml != "")
                {
                    SetTimerDownEnable(50);

                    autoBroswerFrom.currentStep.Text = "货比三家";

                    m_iOtherItemStopTime = autoBroswerFrom.rndGenerator.Next(7, 15);
                    pageExpireTime = DateTime.Now.AddMilliseconds(m_iOtherItemStopTime * millSeconds);
                    isOpenningURL = false;
                }

            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Main)
            {
                SetTimerDownEnable(500);

                if (lastURL != str1)
                {
                    autoBroswerFrom.currentStep.Text = "访问主宝贝";

                    int stopTime = autoBroswerFrom.rndGenerator.Next(m_iMainItemStopMin, m_iMainItemStopMax);

                    string labStr = "主宝贝停留时间:" + stopTime + "S";
                    FileLogger.Instance.LogInfo(labStr);
                    autoBroswerFrom.stopTimeLabel.Text = labStr;

                    isOpenningURL = false;
                    pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
                    getRandClickMainItem();
                }
                
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_MainPage)
            {
                SetTimerDownEnable(300);

                autoBroswerFrom.currentStep.Text = "访问首页";

                int stopTime = 0;
                if (m_isNativeBack)
                {
                    stopTime = autoBroswerFrom.rndGenerator.Next(3, 8);
                }
                else
                {
                    stopTime = autoBroswerFrom.rndGenerator.Next(8, 15);
                }
                string labStr = "首页停留时间:" + stopTime + "S";
                FileLogger.Instance.LogInfo(labStr);
                autoBroswerFrom.stopTimeLabel.Text = labStr;

                isOpenningURL = false;
                pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Other)
            {
                SetTimerDownEnable(500);
                if (lastURL != str1)
                {
                    autoBroswerFrom.currentStep.Text = "访问其它随机宝贝";

                    int stopTime = autoBroswerFrom.rndGenerator.Next(m_iOhterItemStopMin, m_iOtherItemStopMax);
                    string labStr = "其它随机宝贝停留时间:" + stopTime + "S";
                    FileLogger.Instance.LogInfo(labStr);
                    autoBroswerFrom.stopTimeLabel.Text = labStr;

                    isOpenningURL = false;
                    pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
                    getRandClickMainItem();
                }

                
            }
        }
#region 直通车

        ECurrentStep m_lastStep;
        bool bFirstLoadSearch = true;
        public HtmlElement searchIsFoundZTC(ref bool isFound)
        {
            //HtmlElement foundAnchorEle = null;
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("a");
            foreach (HtmlElement linkEle in linkElements)
            {
                // If there's more than one button, you can check the
                //element.InnerHTML to see if it's the one you want
                string titleName = linkEle.GetAttribute("title");
                if (titleName == keyInfo.m_ztcTitle)
                {
                    m_myItemElement = linkEle;
                    break;
                }
            }
            linkElements = InitialTabBrowser.Document.GetElementsByTagName("span");
            foreach (HtmlElement linkEle in linkElements)
            {
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                string className = linkEle.GetAttribute("className");
                if (className == "page-cur")
                {
                    Int32.TryParse(linkEle.InnerText, out currentPage);
                    break;
                }
            }
            if (m_myItemElement == null)
            {
                isFound = false;
                return m_myItemElement;
            }
            isFound = true;
            return m_myItemElement;
        }
        bool FindLinkToClick(HtmlElement root, string text, ref HtmlElement found)
        {
            foreach (var child in root.Children)
            {
                var element = (HtmlElement)child;
                if (element.InnerText == text)
                {
                    found = element.Parent;
                    return true;
                }
                if (FindLinkToClick(element, text, ref found))
                    return true;
            }
            return false;
        }

        public void ClickPrice()
        {
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("input");
            HtmlElement startPriceEle = null;
            HtmlElement endPriceEle = null;
            HtmlElement divBtnPriceEle = null;
            HtmlElement btnPriceEle = null;
            foreach (HtmlElement linkEle in linkElements)
            {
                string className = linkEle.GetAttribute("name");
                if (className == "start_price")
                {
                    startPriceEle = linkEle;
                }
                if (className == "end_price")
                {
                    endPriceEle = linkEle;
                }
            }
            linkElements = InitialTabBrowser.Document.GetElementsByTagName("div");
            foreach (HtmlElement linkEle in linkElements)
            {
                string className = linkEle.GetAttribute("className");
                if (className == "btns clearfix")
                {
                    divBtnPriceEle = linkEle;
                    break;
                }
            }
            if (divBtnPriceEle == null)
            {
                FileLogger.Instance.LogInfo("找不到");
                return;
            }
            InitialTabBrowser.Document.GetElementById("rank-priceform").Focus();
            btnPriceEle = divBtnPriceEle.Children[0].Children[0];
            //Thread.Sleep(1000);
            startPriceEle.ScrollIntoView(true);
            //Rectangle rect = wbElementMouseSimulate.GetElementRect(InitialTabBrowser.Document.Body.DomElement as mshtml.IHTMLElement, endPriceEle.DomElement as mshtml.IHTMLElement);
            //FileLogger.Instance.LogInfo("6:" + p.X); 
            
            //InitialTabBrowser.Document.Window.ScrollTo(0, rect.Top - rect.Height * 2);
            string startPriceStr = keyInfo.m_startPrice.ToString();
            //ClickItemByItem(InitialTabBrowser.Handle, InitialTabBrowser.Document, startPriceEle);
            startPriceEle.SetAttribute("value", startPriceStr);
            //int Y = rect.Top - InitialTabBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop;
            SimulateClick(endPriceEle, new Point(0, 0));
            string endPriceStr = keyInfo.m_endPrice.ToString();
            endPriceEle.SetAttribute("value", endPriceStr);
            endPriceEle.RaiseEvent("onmousedown");
            //string attrName = btnPriceEle.GetAttribute("className");
            //do
            //{
            //    Application.DoEvents();
            //    attrName = btnPriceEle.GetAttribute("className");
            //} while (attrName != "i on");
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            //endPriceEle.Focus();
            //endPriceEle.RaiseEvent("submit");
            //endPriceEle.InvokeMember("click");
            //Thread.Sleep(10000);
            btnPriceEle.SetAttribute("className", "i on");
            ////HtmlElementCollection elements = this.InitialTabBrowser.Document.GetElementsByTagName("Form");

            ////foreach (HtmlElement currentElement in elements)
            ////{
            ////    currentElement.InvokeMember("submit");
            ////}
            //InitialTabBrowser.Document.GetElementById("rank-priceform").Focus();
            //InitialTabBrowser.Document.GetElementById("rank-priceform").InvokeMember("submit");
            //InitialTabBrowser.Document.GetElementById("rank-priceform").All[0].All[1].InvokeMember("submit");
          
            btnPriceEle.Focus();

            //btnPriceEle.InvokeMember("submit");
            //btnPriceEle.RaiseEvent("onclick");
            btnPriceEle.InvokeMember("click", null);
            //btnPriceEle.ScrollIntoView(true);  
            //divBtnPriceEle.All[0].InvokeMember("submit");
            ////Point p = GetOffset(btnPriceEle);
            //SimulateClick(btnPriceEle, new Point(0, 0));
            //mshtml.IHTMLDocument3 doc3 = InitialTabBrowser.Document.DomDocument as mshtml.IHTMLDocument3;
            ////doc3.GET
            //mshtml.IHTMLElement htmlElem = btnPriceEle.DomElement as mshtml.IHTMLElement;
            //htmlElem.click();
        }
        public void doMainJobZTC()
        {
            //bool bRet = false;
            FileLogger.Instance.LogInfo("当前文档状态:" + this.InitialTabBrowser.ReadyState);
            if ((m_currentStep != ECurrentStep.ECurrentStep_Load)
                && this.InitialTabBrowser.ReadyState != WebBrowserReadyState.Complete)
                return;
            //if (m_currentStep == ECurrentStep.ECurrentStep_Search && bFirstLoadSearch &&
            //    InitialTabBrowser.ReadyState != WebBrowserReadyState.Complete)
            //{
            //    return;
            //}
            string innerHtml = this.InitialTabBrowser.Document.Body.InnerHtml;
            string str1 = this.InitialTabBrowser.Document.Url.ToString();
            FileLogger.Instance.LogInfo("当前步骤:" + m_currentStep);
            //FileLogger.Instance.LogInfo("Cookie:" + InitialTabBrowser.Document.Cookie);
            autoBroswerFrom.URLTextBox.Text = str1;

            if (m_currentStep == ECurrentStep.ECurrentStep_Load)
            {
                DateTime dateExpire = DateTime.Parse("2013-12-29 02:30:01");
                if (DateTime.Now > dateExpire)
                {
                    MessageBox.Show("未知错误，可能淘宝又变标签了，请联系作者", "出错啦！");
                    return;
                }
                //bRet = searchBroswer(keyWord);
                if (str1.IndexOf("http://www.taobao.com/") > -1 && innerHtml.IndexOf("淘宝网首页") > -1)
                {
                    this.InitialTabBrowser.Document.GetElementById("q").InnerText = keyInfo.m_keyword;
                    this.InitialTabBrowser.Document.GetElementById("J_TSearchForm").InvokeMember("submit");
                    isOpenningURL = true;
                    m_lastStep = m_currentStep;
                    m_currentStep = ECurrentStep.ECurrentStep_Search;
                    openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    lastURL = str1;
                }
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Search)
            {
                
                if (str1.IndexOf("http://s.taobao.com/search") > -1 && innerHtml.IndexOf("所有分类") > -1 && m_lastStep == ECurrentStep.ECurrentStep_Load)
                {
                    ClickPrice();
                    bFirstLoadSearch = false;
                    m_lastStep = ECurrentStep.ECurrentStep_Search;
                    autoBroswerFrom.currentStep.Text = "查找宝贝";
                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    isOpenningURL = false;
                    lastURL = str1;
                }
                else if (str1.IndexOf("http://s.taobao.com/search") > -1 && innerHtml.IndexOf("所有分类") > -1)
                {
                    //if (str1.Contains("filter=reserve_price") == false)
                    //{
                    //    return;
                    //}
                    SetTimerDownEnable(5);
                    //searchInPage();
                    autoBroswerFrom.currentStep.Text = "查找宝贝";
                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    isOpenningURL = false;
                    lastURL = str1;
                }

            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Compare)
            {
                if (str1.IndexOf("http://s.taobao.com/search") > -1 && innerHtml.IndexOf("所有分类") > -1)
                {
                    SetTimerDownEnable(5);

                    autoBroswerFrom.currentStep.Text = "货比三家回搜索页";

                    pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
                    isOpenningURL = false;
                    //pageMoniterTimer.Enabled = false;
                    //pageMoniterTimer.Stop();
                }
                else if ((str1.IndexOf("http://item.taobao.com") > -1 || str1.IndexOf("http://detail.tmall.com") > -1) && innerHtml != "")
                {
                    SetTimerDownEnable(50);

                    autoBroswerFrom.currentStep.Text = "货比三家";

                    m_iOtherItemStopTime = autoBroswerFrom.rndGenerator.Next(7, 15);
                    pageExpireTime = DateTime.Now.AddMilliseconds(m_iOtherItemStopTime * millSeconds);
                    isOpenningURL = false;
                }

            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Main)
            {
                SetTimerDownEnable(500);

                if (lastURL != str1)
                {
                    autoBroswerFrom.currentStep.Text = "访问主宝贝";

                    int stopTime = autoBroswerFrom.rndGenerator.Next(m_iMainItemStopMin, m_iMainItemStopMax);

                    string labStr = "主宝贝停留时间:" + stopTime + "S";
                    FileLogger.Instance.LogInfo(labStr);
                    autoBroswerFrom.stopTimeLabel.Text = labStr;

                    isOpenningURL = false;
                    pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
                    getRandClickMainItem();
                }

            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_MainPage)
            {
                SetTimerDownEnable(300);

                autoBroswerFrom.currentStep.Text = "访问首页";

                int stopTime = 0;
                if (m_isNativeBack)
                {
                    stopTime = autoBroswerFrom.rndGenerator.Next(3, 8);
                }
                else
                {
                    stopTime = autoBroswerFrom.rndGenerator.Next(8, 15);
                }
                string labStr = "首页停留时间:" + stopTime + "S";
                FileLogger.Instance.LogInfo(labStr);
                autoBroswerFrom.stopTimeLabel.Text = labStr;

                isOpenningURL = false;
                pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Other)
            {
                SetTimerDownEnable(500);
                if (lastURL != str1)
                {
                    autoBroswerFrom.currentStep.Text = "访问其它随机宝贝";

                    int stopTime = autoBroswerFrom.rndGenerator.Next(m_iOhterItemStopMin, m_iOtherItemStopMax);
                    string labStr = "其它随机宝贝停留时间:" + stopTime + "S";
                    FileLogger.Instance.LogInfo(labStr);
                    autoBroswerFrom.stopTimeLabel.Text = labStr;

                    isOpenningURL = false;
                    pageExpireTime = DateTime.Now.AddMilliseconds(stopTime * millSeconds);
                    getRandClickMainItem();
                }


            }
        }
        public bool searchInZTCPage()
        {
            bool isFound = false;
            HtmlElement foundAnchorEle = searchIsFoundZTC(ref isFound);
            if (isFound)
            {
                FileLogger.Instance.LogInfo("在当前页:" + currentPage + " 找到");
                if (m_randCompCount != 0)
                {
                    //randSelectOtherItem();
                    randVisitOtherItemInSearch();
                    m_randCompCount--;
                }
                else
                {
                    visitMe();
                }
            }
            else
            {
                m_currentStep = ECurrentStep.ECurrentStep_Search;
                FileLogger.Instance.LogInfo("在当前页:" + currentPage + ",没有找到，下一页继续");
                if (currentPage >= keyInfo.m_startPage && currentPage <= keyInfo.m_endPage)
                {
                    gotoNextPage();
                }
                else if (currentPage <= keyInfo.m_endPage)
                {
                    //jump to startpage
                    jumpToPage(keyInfo.m_startPage);
                }
                else
                {
                    //stop search
                    jobExpireTime = DateTime.Now.AddMilliseconds(-10 * 1000);
                }

            }
            return true;
        }
        void timeDownZTC_Tick(object sender, EventArgs e)
        {
            HtmlDocument doc = (HtmlDocument)InitialTabBrowser.Document;
            if (doc.Body == null)
            {
                return;
            }
            int height = doc.Body.ScrollRectangle.Height;
            currentScrolBarPos += height / 30;
            if (currentScrolBarPos >= height)
            {
                currentScrolBarPos = height;
                timeDown.Enabled = false;
                timeUp.Enabled = true;

                switch (m_currentStep)
                {
                    case ECurrentStep.ECurrentStep_Search:
                        timeUp.Enabled = false;
                        searchInZTCPage();
                        break;
                }
            }
            doc.Window.ScrollTo(new Point(0, currentScrolBarPos));
        }
        void timeUpZTC_Tick(object sender, EventArgs e)
        {
            HtmlDocument doc = (HtmlDocument)InitialTabBrowser.Document;
            int height = doc.Body.ScrollRectangle.Height;
            currentScrolBarPos -= height / 100;

            if (currentScrolBarPos <= 0)
            {
                currentScrolBarPos = 0;
                timeUp.Enabled = false;
            }
            doc.Window.ScrollTo(new Point(0, currentScrolBarPos));
            if (currentScrolBarPos <= 0)
            {
                switch (m_currentStep)
                {
                    case ECurrentStep.ECurrentStep_Search:
                        //searchInPage();
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Compare:
                        {
                            if (m_isNativeBack)
                            {
                                //又到搜索页
                                if (m_randCompCount == 0)
                                {
                                    bool isFound = false;
                                    HtmlElement foundAnchorEle = searchIsFoundZTC(ref isFound);
                                    if (isFound)
                                    {
                                        visitMe();
                                        m_isNativeBack = false;
                                    }
                                }
                                else
                                {
                                    bool bRandVisitOther = true;
                                    bRandVisitOther = randVisitOtherItemInSearch();
                                    if (bRandVisitOther)
                                    {
                                        m_randCompCount--;
                                        m_isNativeBack = false;
                                    }
                                }

                            }
                        }
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Me_MainPage:
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Me_Main:
                    case ECurrentStep.ECurrentStep_Visit_Me_Other:
                        {
                            if (lastURL != InitialTabBrowser.Document.Url.ToString())
                            {
                                lastURL = InitialTabBrowser.Document.Url.ToString();
                            }
                            clickItemPage();
                        }

                        break;
                }

            }
        }
#endregion
#region 自然点击

        void timeDown_Tick(object sender, EventArgs e)
        {
            HtmlDocument doc = (HtmlDocument)InitialTabBrowser.Document;
            if (doc.Body == null)
            {
                return;
            }
            int height = doc.Body.ScrollRectangle.Height;
            currentScrolBarPos += height / 30;
            if (currentScrolBarPos >= height)
            {
                currentScrolBarPos = height;
                timeDown.Enabled = false;
                timeUp.Enabled = true;

                switch (m_currentStep)
                {
                    case ECurrentStep.ECurrentStep_Search:
                        timeUp.Enabled = false;
                        searchInPage();
                        break;
                }
            }
            doc.Window.ScrollTo(new Point(0, currentScrolBarPos));
        }
        void timeUp_Tick(object sender, EventArgs e)
        {
            HtmlDocument doc = (HtmlDocument)InitialTabBrowser.Document;
            int height = doc.Body.ScrollRectangle.Height;
            currentScrolBarPos -= height / 100;
            
            if (currentScrolBarPos <= 0)
            {
                currentScrolBarPos = 0;
                timeUp.Enabled = false;
            }
            doc.Window.ScrollTo(new Point(0, currentScrolBarPos));
            if (currentScrolBarPos <= 0)
            {
                switch (m_currentStep)
                {
                    case ECurrentStep.ECurrentStep_Search:
                        //searchInPage();
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Compare:
                        {
                            if (m_isNativeBack)
                            {
                                //又到搜索页
                                if (m_randCompCount == 0)
                                {
                                    bool isFound = false;
                                    HtmlElement foundAnchorEle = searchIsFound(ref isFound);
                                    if (isFound)
                                    {
                                        visitMe();
                                        m_isNativeBack = false;
                                    }
                                }
                                else
                                {
                                    bool bRandVisitOther = true;
                                    bRandVisitOther = randVisitOtherItemInSearch();
                                    if (bRandVisitOther)
                                    {
                                        m_randCompCount--;
                                        m_isNativeBack = false;
                                    }
                                }

                            }
                        }
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Me_MainPage:
                        break;
                    case ECurrentStep.ECurrentStep_Visit_Me_Main:
                    case ECurrentStep.ECurrentStep_Visit_Me_Other:
                        {
                            if (lastURL != InitialTabBrowser.Document.Url.ToString())
                            {
                                lastURL = InitialTabBrowser.Document.Url.ToString();
                            }
                            clickItemPage();
                        }
                        
                        break;
                }
                
            }
        }
#endregion
        public bool clickItemPage()
        {
            //详情页面点击
            int alinkCount = m_mainItemClickElement.Count;

            if (alinkCount != 0)
            {
                //点击一次，等下一次到timer-up再点击
                HtmlElement element = m_mainItemClickElement[alinkCount - 1];
                ClickItemByItem(InitialTabBrowser.Handle, InitialTabBrowser.Document, element);
                m_mainItemClickElement.Remove(element);
                timeDown.Enabled = true;
            }
            else
            {
                alinkCount = m_mainItemSpanElement.Count;
                if (alinkCount == 0)
                {
                    RandMove(InitialTabBrowser.Handle, 500, autoBroswerFrom.webBrowserPanel.ClientRectangle);
                    return true;
                }
                //点击一次，等下一次到timer-up再点击
                HtmlElement element = m_mainItemSpanElement[alinkCount - 1];
                ClickItemByItem(InitialTabBrowser.Handle, InitialTabBrowser.Document, element);
                
                m_mainItemSpanElement.Remove(element);
            }
            
            return true;
        }

        public bool getRandClickMainItem()
        {
            //宝贝详情页面 鼠标点击点
            //List<HtmlElement> totalItemLinkList = new List<HtmlElement>();//主宝贝页面里面所有的其它宝贝链接
            List<HtmlElement> totalEnterMainPageLinkList = new List<HtmlElement>();//进入首页的方式
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("a");
            foreach (HtmlElement linkEle in linkElements)
            {
                // If there's more than one button, you can check the
                //element.InnerHTML to see if it's the one you want
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                for (int i = 0; i < m_clickLinkItem.Length; i++ )
                {
                    if (linkEle.InnerText.Contains(m_clickLinkItem[i]))
                    {
                        m_mainItemClickElement.Add(linkEle);
                        break;
                    }
                }

                for (int i = 0; i < m_clickMainPageItem.Length; i++)
                {
                    if (linkEle.InnerText.ToString().Trim() == (m_clickMainPageItem[i]))
                    {
                        totalEnterMainPageLinkList.Add(linkEle);
                        break;
                    }
                }
            }

            int totalMainPageCount = totalEnterMainPageLinkList.Count;
            int enterMainPageIndex = 0;
            if (totalMainPageCount > 1)
            {
                enterMainPageIndex = autoBroswerFrom.rndGenerator.Next(0, totalMainPageCount);
            }

            m_myMainPageElement = totalEnterMainPageLinkList[enterMainPageIndex];

            var spanElements = InitialTabBrowser.Document.GetElementsByTagName("span");
            foreach (HtmlElement linkEle in spanElements)
            {
                // If there's more than one button, you can check the
                //element.InnerHTML to see if it's the one you want
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                for (int i = 0; i < m_clickSpanItem.Length; i++)
                {
                    if (linkEle.InnerText.Contains(m_clickSpanItem[i]))
                    {
                        m_mainItemSpanElement.Add(linkEle);
                        break;
                    }
                }

            }
            return true;
        }
        public bool searchInPage()
        {
            bool isFound = false;
            HtmlElement foundAnchorEle = searchIsFound(ref isFound); 
            if (isFound)
            {
                FileLogger.Instance.LogInfo("在当前页:" + currentPage+ " 找到");
                if (m_randCompCount != 0)
                {
                    //randSelectOtherItem();
                    randVisitOtherItemInSearch();
                    m_randCompCount--;
                }
                else
                {
                    visitMe();
                }
            }
            else
            {
                m_currentStep = ECurrentStep.ECurrentStep_Search;
                FileLogger.Instance.LogInfo("在当前页:" + currentPage + ",没有找到，下一页继续");
                autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 在当前页:" + currentPage + ",没有找到，下一页继续" + "\r\n";
                if (currentPage >= keyInfo.m_startPage && currentPage <= keyInfo.m_endPage)
                {
                    gotoNextPage();
                }
                else if (currentPage <= keyInfo.m_endPage)
                {
                    //jump to startpage
                    jumpToPage(keyInfo.m_startPage);
                }
                else
                {
                    //stop search
                    jobExpireTime = DateTime.Now.AddMilliseconds(-10 * 1000);
                }

            }
            return true;
        }

        public HtmlElement getPicElement(HtmlElement itemBoxEle) //itemBoxEle == <div class=col item>
        {
            //通过itembox 得到pic element
            if (itemBoxEle == null)
            {
                return null;
            }
            HtmlElement divPicEle = itemBoxEle.All[0];
            if (divPicEle == null)
            {
                return null;
            }
            HtmlElement divPicBoxEle = divPicEle.All[0];
            if (divPicBoxEle == null)
            {
                return null;
            }
            return divPicBoxEle;
        }
        public HtmlElement getItemBoxElement(HtmlElement linkEle)
        {
            if (linkEle == null)
            {
                return null;
            }

            HtmlElement parentEle = linkEle.Parent;//"col seller"
            if (parentEle == null)
            {
                return null;
            }

            HtmlElement grandParentEle = parentEle.Parent;//"class="row""
            if (grandParentEle == null)
            {
                return null;
            }
            HtmlElement grandParentEle2 = grandParentEle.Parent;//"item box"
            if (grandParentEle2 == null)
            {
                return null;
            }
            return (HtmlElement)grandParentEle2;
        }
        //搜索页面随机
        public bool randVisitOtherItemInSearch()
        {
            var divCollect = InitialTabBrowser.Document.GetElementsByTagName("div");
            List<HtmlElement> colItemCollect = new List<HtmlElement>();
            foreach (HtmlElement el in divCollect)
            {
                string divClassAttr = el.GetAttribute("className");
                string nidString = el.GetAttribute("nid");
                if (divClassAttr.Trim().StartsWith("col item") && nidString != "")
                {
                    colItemCollect.Add(el);
                }
            }

            //HtmlElementCollection itemChilds = tbContentChildDIV.Children;//item box
            int itemListCount = colItemCollect.Count;
            if (itemListCount == 0)//查找出来的宝贝个数
            {
                return false;
            }

            int randIndex = autoBroswerFrom.rndGenerator.Next(0, itemListCount);

            HtmlElement itemElement = colItemCollect[randIndex];
            HtmlElement picBoxElement = getPicElement(itemElement);

            if (picBoxElement == null)
            {
                FileLogger.Instance.LogInfo("货比三家结束了");
                return false;//访问结束了
            }

            HtmlElement visitItem = picBoxElement;
            visitItem.All[0].All[0].SetAttribute("target", "_top");

            //FileLogger.Instance.LogInfo("开始浏览其他家的" + visitItem.OuterHtml);
            //Tabs.SelectTab(0);//返回 默认的Tab
            //InitialTabBrowser.Document.InvokeScript("eventFire", new object[] { visitItem.All[0].All[0] });     
            ClickItemByPicBox(InitialTabBrowser.Handle, ref visitItem);
            isOpenningURL = true;
            openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
            m_currentStep = ECurrentStep.ECurrentStep_Visit_Compare;
            return true;
        }
        public bool randVisitOtherItem()
        {
            //随机获取其它宝贝
            List<HtmlElement> totalItemLinkList = new List<HtmlElement>();//页面里面所有的其它宝贝链接

            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("a");
            foreach (HtmlElement linkEle in linkElements)
            {
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                string hrefAttrName = linkEle.GetAttribute("href");
                if (otherItemRegex.IsMatch(hrefAttrName) || otherTianMallRegex.IsMatch(hrefAttrName))
                {
                    totalItemLinkList.Add(linkEle);
                }

            }

            //深度随机
            int totalItemCounts = totalItemLinkList.Count;
            FileLogger.Instance.LogInfo("当前访问深度:" + m_randDeepItemCount);
            if (totalItemCounts == 0)
            {
                return false;
            }
            int randItemIndex = autoBroswerFrom.rndGenerator.Next(0, totalItemCounts);
            HtmlElement visitItem = totalItemLinkList[randItemIndex];
            if (visitItem == null)
            {
                return false;//访问结束了
            }
            visitItem.SetAttribute("target", "_top");

            ClickItemByItem(InitialTabBrowser.Handle, InitialTabBrowser.Document, visitItem);
            isOpenningURL = true;
            openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
            m_currentStep = ECurrentStep.ECurrentStep_Visit_Me_Other;
            return true;
        }
        public bool ClickNextPage(IntPtr hwnd, HtmlElement visitItem)
        {
            Point p = GetOffset(visitItem);
            Size winSize = InitialTabBrowser.Document.Window.Size;
            InitialTabBrowser.Document.Window.ScrollTo(winSize.Width / 2, p.Y);
            p.Y -= InitialTabBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop;
            p.X += visitItem.OffsetRectangle.Width / 2;
            p.Y += visitItem.OffsetRectangle.Height / 2;

            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.Parent.DomElement, "click" });                    
            //HtmlElement ele = InitialTabBrowser.Document.GetElementFromPoint(p);
            ClickOnPointInClient(hwnd, p);

            //ClientToScreen(hwnd, ref p);
            //Cursor.Position = new Point(p.X, p.Y);
            visitItem.InvokeMember("click");
            //ClickOnPoint(hwnd, p);

            return true;
        }

        //在搜索页点击
        public bool ClickItemByPicBox(IntPtr hwnd, ref HtmlElement visitItem)
        {
            try
            {
                //if (visitItem == null)
                //{
                //    FileLogger.Instance.LogInfo("不可能");
                //}
                //FileLogger.Instance.LogInfo("1:" + visitItem.InnerHtml);
                Point p = GetOffset(visitItem);
                //FileLogger.Instance.LogInfo("2:" + p.X);
                InitialTabBrowser.Document.Window.ScrollTo(0, p.Y - visitItem.OffsetRectangle.Height / 4);
                //FileLogger.Instance.LogInfo("3:" + p.X);
                p.Y -= InitialTabBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop;
                //FileLogger.Instance.LogInfo("4:" + p.X);
                p.X += visitItem.OffsetRectangle.Width / 4;
                p.Y += visitItem.OffsetRectangle.Height / 4;
                //FileLogger.Instance.LogInfo("5:" + p.X);

                Rectangle rect = wbElementMouseSimulate.GetElementRect(InitialTabBrowser.Document.Body.DomElement as mshtml.IHTMLElement, visitItem.DomElement as mshtml.IHTMLElement);
                //FileLogger.Instance.LogInfo("6:" + p.X);
                rect.Width *= 2;
                rect.Height *= 2;
                //RandMove(hwnd, 500, rect);
                //SimulateClick(visitItem, p);
                ClickOnPointInClient(hwnd, p);
                //ClickOnPoint(hwnd, p);

                return true;
            }
            catch (System.Exception ex)
            {
                FileLogger.Instance.LogInfo("ClickError:" + ex.Message);
                visitItem.All[0].All[0].InvokeMember("click");
            }
            return false;
        }

        //在主页点击其它标记
        public bool ClickItemByItem(IntPtr hwnd, HtmlDocument doc, HtmlElement visitItem)
        {
            Point p = GetOffset(visitItem);
            doc.Window.ScrollTo(0, p.Y);
            p.Y -= doc.GetElementsByTagName("HTML")[0].ScrollTop;

            Rectangle rect = wbElementMouseSimulate.GetElementRect(doc.Body.DomElement as mshtml.IHTMLElement, visitItem.DomElement as mshtml.IHTMLElement);
            p.X = rect.Left + 3;
            p.Y = rect.Top + 3;
            
            RandMove(hwnd, 500, rect);
            //SimulateClick(visitItem, p);
            ClickOnPointInClient(hwnd, p);
            //ClientToScreen(hwnd, ref p);
            //Cursor.Position = p;
            return true;
        }

        public bool visitMe()
        {
            if (m_myItemElement == null)
            {
                return false;
            }

            HtmlElement itemBoxEle = getItemBoxElement(m_myItemElement);
            if (itemBoxEle == null)
            {
                return false;
            }

            HtmlElement itemPicEle = itemBoxEle.All[0].All[0];
            itemPicEle.All[0].SetAttribute("target", "_top");
            ClickItemByPicBox(InitialTabBrowser.Handle, ref itemPicEle);
            isOpenningURL = true;
            openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
            m_currentStep = ECurrentStep.ECurrentStep_Visit_Me_Main;
            return true;
        }
        public bool enterMainPage()
        {
            if (m_myMainPageElement == null)
            {
                return false;
            }
            m_myMainPageElement.SetAttribute("target", "_top");
            FileLogger.Instance.LogInfo("进入首页:" + m_myMainPageElement.InnerHtml);
            ClickItemByItem(InitialTabBrowser.Handle, InitialTabBrowser.Document, m_myMainPageElement);
            isOpenningURL = true;
            openURLExpireTime = DateTime.Now.AddMilliseconds(OPENURLTIMEOUT);
            m_currentStep = ECurrentStep.ECurrentStep_Visit_Me_MainPage;
            return true;
        }

        public HtmlElement searchIsFound(ref bool isFound)
        {
            //HtmlElement foundAnchorEle = null;
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("a");
            foreach (HtmlElement linkEle in linkElements)
            {
                // If there's more than one button, you can check the
                //element.InnerHTML to see if it's the one you want
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                if (linkEle.InnerText.ToString().Trim() == (autoBroswerFrom.getSellerName()) && 
                    linkEle.Parent.TagName.ToLower().Trim() == "div")
                {
                    m_myItemElement = linkEle;
                    break;
                }
            }
            linkElements = InitialTabBrowser.Document.GetElementsByTagName("span");
            foreach (HtmlElement linkEle in linkElements)
            {
                if (linkEle.InnerText == null)
                {
                    continue;
                }
                string className = linkEle.GetAttribute("className");
                if (className == "page-cur")
                {
                    Int32.TryParse(linkEle.InnerText, out currentPage);
                    break;
                }
            }
            if (m_myItemElement == null)
            {
                isFound = false;
                return m_myItemElement;
            }
            isFound = true;
            return m_myItemElement;
        }

        public string pageInfo;
        public string prevPage;
        public string nextPage = "1/100";
        public int currentPage = 0;
        //public bool gotoNextPage()
        //{
        //    HtmlElement foundAnchorEle = null;
        //    var linkElements = InitialTabBrowser.Document.GetElementsByTagName("div");
        //    foreach (HtmlElement linkEle in linkElements)
        //    {
        //        string className = linkEle.GetAttribute("className");
        //        if (className == "page-top")
        //        {
        //            pageInfo = linkEle.InnerText;
        //            foundAnchorEle = linkEle;
        //            break;
        //        }
        //    }
        //    HtmlElement nextPageLink = null;
        //    if (foundAnchorEle == null)
        //    {
        //        return false;
        //    }
        //    foreach (HtmlElement linkEle in foundAnchorEle.All)
        //    {
        //        string className = linkEle.GetAttribute("className");
        //        if (className == "page-next")
        //        {
        //            HtmlElement pageInfoEle = linkEle.FirstChild;
        //            //pageInfo = pageInfoEle.InnerText;
        //            nextPageLink = linkEle;
        //            break;
        //        }
        //    }
        //    if (nextPage == null || nextPageLink == null)
        //    {
        //        FileLogger.Instance.LogInfo("没有找到宝贝！");
        //        isNormalQuit = true;
        //        ShutDownWinForms();
        //        return false;
        //    }

        //    nextPageLink = nextPageLink.All[0];
        //    //nextPageLink.InvokeMember("click");
        //    ClickNextPage(InitialTabBrowser.Handle, nextPageLink);
        //    //nextPageLink.InvokeMember("click");//.click();
        //    return true;
        //}
        public bool gotoNextPage()
        {
            HtmlElement foundAnchorEle = null;
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("li");
            foreach (HtmlElement linkEle in linkElements)
            {
                string className = linkEle.GetAttribute("className");
                if (className == "next show")
                {
                    pageInfo = linkEle.InnerText;
                    foundAnchorEle = linkEle;
                    break;
                }
            }
            HtmlElement nextPageLink = null;
            if (foundAnchorEle == null)
            {
                return false;
            }
            nextPageLink = foundAnchorEle.All[0];
            if (nextPage == null || nextPageLink == null)
            {
                FileLogger.Instance.LogInfo("没有找到宝贝！");
                isNormalQuit = true;
                ShutDownWinForms();
                return false;
            }

            
            //nextPageLink.InvokeMember("click");
            ClickNextPage(InitialTabBrowser.Handle, nextPageLink);
            //nextPageLink.InvokeMember("click");//.click();
            return true;
        }
        public bool jumpToPage(int pageIndex)
        {
            HtmlElement inputPageEle = null;
            HtmlElement btnJumpEle = null;
            var linkElements = InitialTabBrowser.Document.GetElementsByTagName("input");
            foreach (HtmlElement linkEle in linkElements)
            {
                string className = linkEle.GetAttribute("className");
                if (className.ToLower() == "page-num")
                {
                    inputPageEle = linkEle;
                    break;
                }
            }
            HtmlElement nextPageLink = null;
            if (inputPageEle == null)
            {
                return false;
            }
            string pageIndexStr = pageIndex.ToString();
            inputPageEle.SetAttribute("value", pageIndexStr);

            linkElements = InitialTabBrowser.Document.GetElementsByTagName("a");
            foreach (HtmlElement linkEle in linkElements)
            {
                string className = linkEle.GetAttribute("className");
                if (className == "btn-jump")
                {
                    btnJumpEle = linkEle;
                    break;
                }
            }
            if (btnJumpEle == null)
            {
                return false;
            }
            btnJumpEle.InvokeMember("click");//.click();
            return true;
        }
        public Point GetOffset(HtmlElement el)
        {
            //get element pos
            Point pos = new Point(el.OffsetRectangle.Left, el.OffsetRectangle.Top);

            //get the parents pos
            HtmlElement tempEl = el.OffsetParent;
            while (tempEl != null)
            {
                pos.X += tempEl.OffsetRectangle.Left;
                pos.Y += tempEl.OffsetRectangle.Top;
                tempEl = tempEl.OffsetParent;
            }

            return pos;
        }

        private void RandMove(IntPtr wndHandle, int moveTime, Rectangle randMoveRect)
        {
            int currentRandMoveTimes = 0;
            int randMoveCount = moveTime / randMoveInterval;
            BlockInput(true);
            while (currentRandMoveTimes <= randMoveCount)
            {
                int randOffsetX = autoBroswerFrom.rndGenerator.Next(0, randMoveRect.Width);
                int randOffsetY = autoBroswerFrom.rndGenerator.Next(0, randMoveRect.Height);
                Point clientPoint = new Point(randOffsetX + randMoveRect.Left, randOffsetY + randMoveRect.Top);

                ClientToScreen(wndHandle, ref clientPoint);
                /// set cursor on coords, and press mouse
                //BlockInput(true);
                Cursor.Position = new Point(clientPoint.X, clientPoint.Y);
                currentRandMoveTimes++;
                Thread.Sleep(randMoveInterval);
            }

            BlockInput(false);
        }

        private void ClickOnPointInClient(IntPtr wndHandle, Point clientPoint)
        {
            int position = ((clientPoint.Y) << 16) | (clientPoint.X );

            IntPtr handle = wndHandle;
            StringBuilder className = new StringBuilder(100);
            while (className.ToString() != "Internet Explorer_Server") // your mileage may vary with this classname
            {
                handle = GetWindow(handle, 5); // 5 == child
                GetClassName(handle, className, className.Capacity);
            }
            const UInt32 WM_LBUTTONDOWN = 0x0201;
            const UInt32 WM_LBUTTONUP = 0x0202;
            const int MK_LBUTTON = 0x0001;
            // 模拟鼠标按下  
            FileLogger.Instance.LogInfo("before button down");
            AutoBroswerForm.SendMessage(handle, WM_LBUTTONDOWN, MK_LBUTTON, position);
            FileLogger.Instance.LogInfo("after button down");
            AutoBroswerForm.SendMessage(handle, WM_LBUTTONUP, MK_LBUTTON, position);
            FileLogger.Instance.LogInfo("after button up");
        }

        private void SimulateClick(HtmlElement visitItem, Point p)
        {
            if (visitItem == null)
            {
                FileLogger.Instance.LogInfo("error simulate click");
                return;
            }
            p.X = p.Y = 0;
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "focus", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "mouseover", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "mousemove", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "mouseout", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "mousedown", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "mouseup", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" });
            InitialTabBrowser.Document.InvokeScript("simulate", new object[] { visitItem.DomElement, "click", "{ pointerX: " + p.X + ", pointerY: " + p.Y + " }" }); 

        }
        public void SetTimerUpEnable(int tickInter)
        {
            timeDown.Enabled = false;
            timeDown.Stop();

            timeUp.Interval = tickInter;
            timeUp.Enabled = true;
            timeUp.Start();
        }

        public void SetTimerDownEnable(int tickInter)
        {
            timeUp.Interval = tickInter;
            timeUp.Enabled = false;
            timeUp.Stop();

            timeDown.Interval = tickInter;
            timeDown.Enabled = true;
            timeDown.Start();
        }
        public int passJobTime = 0;
        public void TimerTick(object source, EventArgs e)
        {
            passJobTime += 1;
            autoBroswerFrom.JobPassTimeLabel.Text = passJobTime.ToString() + "S";
            if (isOpenningURL && DateTime.Now > openURLExpireTime)
            {
                FileLogger.Instance.LogInfo("CurrentState:" + m_currentStep + "browserSState:"+InitialTabBrowser.ReadyState);
                FileLogger.Instance.LogInfo("OpenURL fail:" + InitialTabBrowser.Document.Url.ToString());
                autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  OpenURL fail:" + InitialTabBrowser.Document.Url.ToString() + "\n";
                isNormalQuit = true;
                ShutDownWinForms();
            }
            if (DateTime.Now > jobExpireTime)
            {
                FileLogger.Instance.LogInfo("CurrentState:" + m_currentStep);
                FileLogger.Instance.LogInfo("任务超时了:" + InitialTabBrowser.Document.Url.ToString());
                autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  任务超时了:" + InitialTabBrowser.Document.Url.ToString() + " 任务时间:" + jobExpireTime.ToString("yyyy-MM-dd HH:mm:ss") + "\n";
                isNormalQuit = true;
                ShutDownWinForms();
            }
            if (DateTime.Now > pageExpireTime)
            {
                PageMoniterTimeEvent( source, e);
            }
        }
        private bool m_isNativeBack = false;
        public void PageMoniterTimeEvent( object source, EventArgs e)
        {
            timeUp.Enabled = false;
            timeDown.Enabled = false;

            passJobTime = 0;
            if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Compare)
            {
                if (InitialTabBrowser.CanGoBack &&　m_isNativeBack == false)
                {
                    InitialTabBrowser.GoBack();
                    m_isNativeBack = true;
                }
                //pageMoniterTimer.Enabled = false;
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Main)
            {
                enterMainPage();
                FileLogger.Instance.LogInfo("访问主宝贝结束了,进入宝贝首页");
                autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  访问主宝贝结束了,进入宝贝首页" + "\n";

            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_MainPage)
            {
                bool bRandVisitOther = true;
                bRandVisitOther = randVisitOtherItem();
                if (bRandVisitOther == false)//has visit done
                {
                    FileLogger.Instance.LogInfo("首页访问结束，找不到其它宝贝？");
                    autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  首页访问结束，找不到其它宝贝？" + "\n";
                    isNormalQuit = true;
                    ShutDownWinForms();
                }
            }
            else if (m_currentStep == ECurrentStep.ECurrentStep_Visit_Me_Other)
            {

                if (m_randDeepItemCount == 0)
                {
                    //job done
                    FileLogger.Instance.LogInfo("随机宝贝访问结束，找不到其它宝贝？");
                    autoBroswerFrom.LogInfoTextBox.Text += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  随机宝贝访问结束，找不到其它宝贝？" + "\n";
                    isNormalQuit = true;
                    ShutDownWinForms();
                }
                else
                {
                    bool bRandVisitOther = true;
                    bRandVisitOther = randVisitOtherItem();
                    if (bRandVisitOther)
                    {
                        m_randDeepItemCount--;
                    }
                }
                
            }
            pageExpireTime = DateTime.Now.AddMilliseconds(ImpossibleTime);
        }


        public void ShutDownWinForms()
        {
            timeUp.Enabled = false;
            timeUp.Stop();
            timeDown.Enabled = false;
            timeDown.Stop();
            moniterTimer.Enabled = false;
            moniterTimer.Stop();
            if (autoBroswerFrom.isDebugCBChecked() == false)
            {
                this.Close();
                this.Dispose();
            }
            else
            {
                MessageBox.Show("任务完成了", "提示");
            }
        }
    }
}

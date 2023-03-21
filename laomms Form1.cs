using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;



namespace CallSkypeClient
{

    public partial class Form1 : Form
    {
        private static ChromeDriver driver;
        private static ChromeDriverService service;
        private bool finished = false;
        private bool GetOneOrTwo = false;
        private bool Verification = false;
        private bool WindowsOffice = false;
        private bool SelectVersion = false;
        private bool continueActivation = false;
        private bool haveProductkey = false;
        private bool haveErrorcode = false;
        private bool IIDWizard = false;
        private bool IID_Group1 = false;

        private bool ErrorIID = false;
        private bool numberOfWindows = false;
        private bool unablValidate = false;
        private bool getConfirmationId = false;
        private bool finishedGetCID = false;
        private string[] idChunks;
        private readonly ILogger _logger;

        public object Dubeg { get; private set; }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        public Form1()
        {
            InitializeComponent();
        }
        private  void WinForm1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (driver != null )
            {
                driver.Quit ();
                service.Dispose();
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
           
          
            var services = new ServiceCollection();
            services.AddSingleton<RichTextBox>(richTextBox1);
            services.AddLogging(logging =>
            {
                logging.AddProvider(new RichTextBoxLoggerProvider(richTextBox1, (category, level) => level >= Microsoft.Extensions.Logging.LogLevel.Information));
            });

#if DEBUG
            services.Configure<LoggerFilterOptions>(options => options.MinLevel = Microsoft.Extensions.Logging.LogLevel.Debug);
#endif



          
        }


        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan elapsed = DateTime.Now - new DateTime(2023, 1, 1, 0, 0, 0);
            string elapsedTime = elapsed.ToString("hh\\:mm\\:ss"); 


            timer_label.Invoke(new Action(() => {
                timer_label.Text = elapsedTime;
            }));
        }
       
        private void button1_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox2.Text)) return;
            richTextBox1.Text = "";

            if (textBox3.Text.Length == 54 || textBox3.Text.Length == 63)
            {
                int chunkSize = textBox3.Text.Length / 9;
                idChunks = Enumerable.Range(0, 9)
                    .Select(i => textBox3.Text.Substring(i * chunkSize, chunkSize))
                    .ToArray();

            }
            else
                return;

             finished = false;
         GetOneOrTwo = false;
         Verification = false;
         WindowsOffice = false;
         SelectVersion = false;
         continueActivation = false;
         haveProductkey = false;
         haveErrorcode = false;
         IIDWizard = false;
         IID_Group1 = false;
            ErrorIID = false;
            unablValidate = false;
            numberOfWindows = false;
            getConfirmationId = false;
         finishedGetCID = false;
        button1.Enabled = false;
            var startTime = DateTime.Now;
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000; // 1 秒钟
            timer.Tick += (sender, args) =>
            {
                // 计时器每秒钟更新一次文本
                var timeSpan = DateTime.Now - startTime;
                timer_label.Text = timeSpan.ToString(@"hh\:mm\:ss");
            };
            timer.Start();



            service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true; // 隐藏控制台窗口

            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--headless");
            //options.AddArgument("--window-position=-32000,-32000");
            options.AddArgument("--window-size=500,600");
            options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 1);
            options.AddArgument("--allow-all-secure-origins");
            options.AddArgument("--disable-notifications");
            options.AddArgument("--use-fake-ui-for-media-stream");
            options.AddArgument("--mute-audio");
            // 禁用最大化窗口
            options.AddArgument("--start-maximized=false");

            // 禁用快捷键
            //options.AddArgument("--disable-extensions");
            //options.AddArgument("--disable-infobars");
            //options.AddArgument("--disable-popup-blocking");
            //options.AddArgument("--disable-save-password-bubble");
            // 禁用菜单栏和工具栏
            options.AddArgument("--disable-blink-features=RootLayerScrolling");
            // 禁用全屏模式
            options.AddArgument("--disable-fullscreen");


            // 创建 ChromeDriver 对象
            driver = new ChromeDriver(service, options);

          
            Thread runRhread = new Thread(() => LoadWeb(textBox1.Text, textBox2.Text));
            runRhread.Start();



        }
 
        private void LoadWeb(string username,string password)
        {

            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionLength = 0;
                richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "准备登录..."}{Environment.NewLine}");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            });

            driver.Navigate().GoToUrl("https://web.skype.com");
            Thread.Sleep(500);

            // 找到用户名输入框
            IWebElement usernameInput = driver.FindElement(By.Id("i0116"));

            // 输入用户名
            usernameInput.SendKeys(username);

            // 找到下一页按钮并点击
            IWebElement nextPage = driver.FindElement(By.Id("idSIButton9"));
            nextPage.Click();

            Thread.Sleep(1000);

            // 找到密码输入框
            IWebElement passwordInput = driver.FindElement(By.Id("i0118"));

            // 输入密码
            passwordInput.SendKeys(password);

            // 找到登录按钮并点击
            IWebElement login = driver.FindElement(By.Id("idSIButton9"));
            login.Click();

            Thread.Sleep(2000);
            // 找到保持登录按钮并点击
            IWebElement StaySignedIn = driver.FindElement(By.Id("idSIButton9"));
            StaySignedIn.Click();

            // 等待页面加载完毕
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(d => d.Title.StartsWith("Skype"));

            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionLength = 0;
                richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "登录完成"}{Environment.NewLine}");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            });


            // 找到拨号面板按钮
            IWebElement UseDialPad = new WebDriverWait(driver, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div.css-1dbjc4n button[title='Use dial pad'], div.css-1dbjc4n button[title='使用拨号盘']")));
            UseDialPad.Click();

            Thread.Sleep(2000);
            // 在搜索框中输入号码
            IWebElement inputElement = driver.FindElement(By.XPath("//input[@placeholder='Phone number' or @placeholder='电话号码']"));

            inputElement.SendKeys("+448000188354");
            inputElement.SendKeys(OpenQA.Selenium.Keys.Enter);

            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionLength = 0;
                richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "正在拨打电话..."}{Environment.NewLine}");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            });

            Thread runThread = new Thread(() =>
            {
                try
                {
                    var Callfailed = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[contains(text(), 'is unavailable.')]")));
                    richTextBox1.Invoke((MethodInvoker)delegate
                    {
                        richTextBox1.SelectionStart = richTextBox1.TextLength;
                        richTextBox1.SelectionLength = 0;
                        richTextBox1.AppendText($"[{DateTime.Now.ToString()}] {"打电话失败!"}{Environment.NewLine}");
                        richTextBox1.SelectionColor = richTextBox1.ForeColor;
                    });
                    driver.Quit();
                    service.Dispose();
                    return;
                }
                catch (WebDriverTimeoutException ex)
                {
                    Console.WriteLine($"未能找到打电话失败内容：{ex.Message}");
                }
            });
            runThread.Start();




            //翻译按钮
            try
            {
                var translateButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div[data-text-as-pseudo-element='Translate this call'], div[data-text-as-pseudo-element='翻译此调用']")));
                translateButton.Click();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                IWebElement parentDiv = driver.FindElement(By.XPath("//div[@class='css-1dbjc4n r-18u37iz']/parent::div"));
                IWebElement firstChildDiv = parentDiv.FindElement(By.XPath("./div[1]"));
                string language= firstChildDiv.Text;                
                parentDiv.Click();
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "目前翻译情况：" + language}{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });

                if (language.Split('-')[0].Trim() != "英语(美国)" && language.Split('-')[0].Trim() != "English (US)")
                {
                    richTextBox1.Invoke((MethodInvoker)delegate
                    {
                        richTextBox1.SelectionStart = richTextBox1.TextLength;
                        richTextBox1.SelectionLength = 0;
                        richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "修改初始语言为：英语(美国)" }{Environment.NewLine}");
                        richTextBox1.SelectionColor = richTextBox1.ForeColor;
                    });
                    IWebElement language1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + language.Split('-')[0].Trim() + "']"));
                    language1.Click();
                    Thread.Sleep(200);
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    IWebElement languageButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[@aria-label='英语(美国)'] | //button[@aria-label='English (US)']")));
                    languageButton.Click();
                }
                if (language.Split('-')[1].Trim() != "英语(美国)" && language.Split('-')[0].Trim() != "English (US)")
                {
                    richTextBox1.Invoke((MethodInvoker)delegate
                    {
                        richTextBox1.SelectionStart = richTextBox1.TextLength;
                        richTextBox1.SelectionLength = 0;
                        richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "修改目标语言为：英语(美国)" }{Environment.NewLine}");
                        richTextBox1.SelectionColor = richTextBox1.ForeColor;
                    });
                    IWebElement language2 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + language.Split('-')[1].Trim() + "']"));
                    language2.Click();
                    Thread.Sleep(200);
                    IWebElement languageButton = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[@aria-label='英语(美国)'] | //button[@aria-label='English (US)']")));
                    languageButton.Click();

                }
                IWebElement buttonDone = driver.FindElement(By.XPath("//button[@title='Done']"));
                buttonDone.Click();
            }
            catch (WebDriverTimeoutException ex)
            {
                Console.WriteLine($"未能找到翻译按钮：{ex.Message}");
            }

            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionLength = 0;
                richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "电话已接通..."}{Environment.NewLine}");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            });

            //开始通话的标志，拨号面板可点击
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            var element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//button[contains(@title, 'Enter numbers using the dial pad') or @title='使用拨号盘输入号码']")));



            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionLength = 0;
                richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "开始监听内容:"}{Environment.NewLine}");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            });

            // 获取第一次加载的所有 data-text-as-pseudo-element 属性的 div 元素的文本
            var divElements = driver.FindElements(By.CssSelector("div[data-text-as-pseudo-element]:not([data-text-as-pseudo-element=''])"));
            foreach (var divElement in divElements)
            {
                Debug.Print(divElement.Text);
            }


            //开始循环取翻译内容
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            Thread thread = new Thread(() =>
            {
                while (!finished)
                {
                    //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    //string pageSource = (string)js.ExecuteScript("return document.documentElement.outerHTML;");
                    //Regex regex = new Regex(@"<div class=""css-1dbjc4n"">([\s\S]*?)</div>");
                    //MatchCollection matchesDiv = regex.Matches(pageSource);
                    //if (matchesDiv.Count > 0)
                    //{
                    //    MatchCollection matches = Regex.Matches(matchesDiv[0].Groups[1].Value, @"<div\s+data-text-as-pseudo-element\s*=\s*""([^""]*)"""); // @"<div\s+data-text-as-pseudo-element\s*=\s*""([^""]*)"""); 
                    //    List<string> uniqueMatches = matches.Cast<Match>().Select(match => match.Groups[1].Value).Distinct().ToList();
                    //    foreach (string match in uniqueMatches)
                    //    {
                    //        Debug.Print(match);
                    //    }
                    //}


                    ReadOnlyCollection<IWebElement> elements = driver.FindElements(By.CssSelector("div[data-text-as-pseudo-element]"));
                    foreach (IWebElement element in elements)
                    {
                        string script = "return window.getComputedStyle(arguments[0], '::after').getPropertyValue('content')";
                        try
                        {
                            string content = (string)jsExecutor.ExecuteScript(script, element);
                            Debug.WriteLine(content);
                        }
                        catch
                        {

                        }
                        
                    }       


                }
            });
            thread.Start();

           


        }


        private  void richTextBox1_TextChanged(object sender, EventArgs e)
        {


            string VerificationCode = string.Empty;
 
            
            Match match = Regex.Match(richTextBox1.Text, @"\b(\d{3})\b");
            if (richTextBox1.Text.ToLower().Contains ("press one or") &&  !GetOneOrTwo)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "正在拨按键1..."}{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });

               
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
                var button = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath("//button[contains(@title, 'Enter numbers using the dial pad') or @title='使用拨号盘输入号码']")));

                Actions actions = new Actions(driver);
                    actions.MoveToElement(button).Click().Perform();

                    //拨号面板拨1
                    IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                    actions = new Actions(driver);
                    actions.MoveToElement(button1).Click().Perform();
                    GetOneOrTwo = true;
                
            }
            else if (match.Success && !Verification  && richTextBox1.Text.ToLower().Contains("security") && richTextBox1.Text.ToLower().Contains("purposes") && !Verification  || richTextBox1.Text.ToLower().Contains("let's try once more for security purposes" ) && !Verification)
            {
                VerificationCode = match.Value; 
           
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "检测到验证码：" + VerificationCode}{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });

                    IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + VerificationCode.Substring(0, 1) + "']"));
                    button1.Click();

                    Thread.Sleep(200);
                    IWebElement button2 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + VerificationCode.Substring(1, 1) + "']"));
                    button2.Click();

                    Thread.Sleep(200);
                    IWebElement button3 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + VerificationCode.Substring(2, 1) + "']"));
                   button3.Click();


                VerificationCode = string.Empty;
                    Verification = true;
                
            }

            else if ( richTextBox1.Text.ToLower().Contains("didn't") && richTextBox1.Text.ToLower().Contains("get") && richTextBox1.Text.ToLower().Contains("that") && !Verification)
            {

                match = Regex.Match(richTextBox1.Text, @"(?<=\[\d{2}/\d{2}/\d{4}\s\d{2}:\d{2}:\d{2}\]\s)\d{2}\.");
                if (match.Success )
                {
                    VerificationCode = match.Value;

                    IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + VerificationCode.Substring(0, 1) + "']"));
                        Actions actions = new Actions(driver);
                        actions.MoveToElement(button1).Click().Perform();

                    Thread.Sleep(100);
                    IWebElement button2 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + VerificationCode.Substring(1, 1) + "']"));
                        actions.MoveToElement(button2).Click().Perform();

                        VerificationCode = string.Empty;
                        Verification = true;
                    richTextBox1.Invoke((MethodInvoker)delegate
                    {
                        richTextBox1.SelectionStart = richTextBox1.TextLength;
                        richTextBox1.SelectionLength = 0;
                        richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "检测到验证码：" + VerificationCode}{Environment.NewLine}");
                        richTextBox1.SelectionColor = richTextBox1.ForeColor;
                    });
                }
            }

            else if ( richTextBox1.Text.ToLower().Contains("windows") && richTextBox1.Text.ToLower().Contains("office") && richTextBox1.Text.ToLower().Contains("mac") && !WindowsOffice)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { (radioButton1.Checked ? "选择为Windows的安装ID" : "选择为Office的安装ID")  }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });

                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='"+ (radioButton1.Checked ? "1" : "2") + "']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                WindowsOffice = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("if you have previously upgraded")  && richTextBox1.Text.ToLower().Contains("otherwise press") && !SelectVersion)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "试图激活项选择按键1" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                SelectVersion = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("get started") && richTextBox1.Text.ToLower().Contains("continue with phone system")  && !continueActivation)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "按1开始" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                continueActivation = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("product key to activate")  && !haveProductkey)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "没有密钥，按键2" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='2']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                haveProductkey = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("you have that error code handy") && !haveErrorcode)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "没有错误代码，按键2" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='2']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                haveErrorcode = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("computer") && richTextBox1.Text.ToLower().Contains("activation") && richTextBox1.Text.ToLower().Contains("wizard") && !IIDWizard ||  richTextBox1.Text.ToLower().Contains("screen") && richTextBox1.Text.ToLower().Contains("open") && !IIDWizard)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "是否打开电脑屏幕选择按键1" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                IIDWizard = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("sorry") && richTextBox1.Text.ToLower().Contains("didn't seem") && richTextBox1.Text.ToLower().Contains("valid group") && !IIDWizard && !ErrorIID)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "安装ID错误!" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                textBox4.Invoke((MethodInvoker)delegate
                {
                    textBox4.Text = "安装ID错误!";
                });
                ErrorIID = true;
                IWebElement closeButton = driver.FindElement(By.CssSelector("button[title='Close'], button[aria-label='Close']"));
                closeButton.Click();
                IWebElement EndCall = driver.FindElement(By.CssSelector("button[aria-label='End Call'], button[aria-label='结束通话']"));
                EndCall.Click();
                button1.Invoke((MethodInvoker)delegate
                {
                    button1.Enabled = true;

                });
                finished = true;
                
            }
            else if (richTextBox1.Text.ToLower().Contains("installation id") && richTextBox1.Text.ToLower().Contains("group one") && !IID_Group1 || richTextBox1.Text.ToLower().Contains("installation id") && richTextBox1.Text.ToLower().Contains("group 1") && ! IID_Group1 || richTextBox1.Text.ToLower().Contains("enter the first group of digits") && !IID_Group1)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "输入第一组安装ID：" + idChunks[0]}{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                InputIID(idChunks[0]);
                IID_Group1 = true;
                Thread.Sleep(500);
                InputIID(idChunks[1]);
                Thread.Sleep(500);
                InputIID(idChunks[2]);
                Thread.Sleep(500);
                InputIID(idChunks[3]);
                Thread.Sleep(500);
                InputIID(idChunks[4]);
                Thread.Sleep(500);
                InputIID(idChunks[5]);
                Thread.Sleep(500);
                InputIID(idChunks[6]);
                Thread.Sleep(500);
                InputIID(idChunks[7]);
                Thread.Sleep(500);
                InputIID(idChunks[8]);
            }
            
            else if (richTextBox1.Text.ToLower().Contains("installed with") && richTextBox1.Text.ToLower().Contains("copy of") && richTextBox1.Text.ToLower().Contains("different computers") && !numberOfWindows)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "输入Widnows拷贝数量，按键1" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                numberOfWindows = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("unable") && richTextBox1.Text.ToLower().Contains("validate") && richTextBox1.Text.ToLower().Contains("installation") && !unablValidate)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "获取失败!" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                textBox4.Invoke((MethodInvoker)delegate
                {
                    textBox4.Text = "获取失败!";
                });
                unablValidate = true;
                IWebElement closeButton = driver.FindElement(By.CssSelector("button[title='Close'], button[aria-label='Close']"));
                closeButton.Click();
                IWebElement EndCall = driver.FindElement(By.CssSelector("button[aria-label='End Call'], button[aria-label='结束通话']"));
                EndCall.Click();
                button1.Invoke((MethodInvoker)delegate
                {
                    button1.Enabled = true;

                });
                finished = true;
               
            }
            else if (richTextBox1.Text.ToLower().Contains("confirmation id") && richTextBox1.Text.ToLower().Contains("you're ready to get started") && richTextBox1.Text.ToLower().Contains("need a few more moments") && !getConfirmationId)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "准备报读确认ID，准备就绪，按键1" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='1']"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(button1).Click().Perform();
                haveErrorcode = true;
                getConfirmationId = true;
            }
            else if (richTextBox1.Text.ToLower().Contains("repeat that final block") && richTextBox1.Text.ToLower().Contains("you're finished") && !finishedGetCID)
            {
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "报读完毕" }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });

                Regex regex = new Regex(@"Block ([a-zA-Z]+) is (\d{6})");
                MatchCollection matches = regex.Matches(richTextBox1.Text);
                HashSet<string> results = new HashSet<string>();
                foreach (Match matchItem in matches)
                {
                    string blockName = matchItem.Groups[1].Value;
                    string blockValue = matchItem.Groups[2].Value;
                    results.Add(blockValue);
                }
                string output = string.Join("-", results.ToArray());
                richTextBox1.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionLength = 0;
                    richTextBox1.AppendText($"[{DateTime.Now.ToString()}] { "获取CID完毕:" + output }{Environment.NewLine}");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                });
                textBox4.Invoke((MethodInvoker)delegate
                {
                    textBox4.Text = output;
                });
                finishedGetCID = true;
                IWebElement closeButton = driver.FindElement(By.CssSelector("button[title='Close'], button[aria-label='Close']"));
                closeButton.Click();
                IWebElement EndCall = driver.FindElement(By.CssSelector("button[aria-label='End Call'], button[aria-label='结束通话']"));
                EndCall.Click();
                button1.Invoke((MethodInvoker)delegate
                {
                    button1.Enabled = true;

                });
                
                finished = true;
            }
        }

        private void InputIID(string IID)
        {
            IWebElement button1 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(0, 1) + "']"));
            button1.Click();

            Thread.Sleep(100);
            IWebElement button2 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(1, 1) + "']"));
            button2.Click();

            Thread.Sleep(100);
            IWebElement button3 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(2, 1) + "']"));
            button3.Click();

            Thread.Sleep(100);
            IWebElement button4 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(3, 1) + "']"));
            button4.Click();

            Thread.Sleep(100);
            IWebElement button5 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(4, 1) + "']"));
            button5.Click();

            Thread.Sleep(100);
            IWebElement button6 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(5, 1) + "']"));
            button6.Click();

            if (IID.Length == 7)
            {
                Thread.Sleep(100);
                IWebElement button7 = driver.FindElement(By.CssSelector("div[data-text-as-pseudo-element='" + IID.Substring(6, 1) + "']"));
                button7.Click();
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string InstalltionID = System.Text.RegularExpressions.Regex.Replace(textBox3.Text, "[^\\d]", "");
            if (InstalltionID.Trim().Length == 54 || InstalltionID.Trim().Length == 63)
            {
                textBox3.Text = InstalltionID;
                button1.Enabled = true;
                
            }
            else
            {
                button1.Enabled = false;
               
            }
        }
        private void textBox3_Click(object sender, EventArgs e)
        {
            string InstalltionID = null;
            string clipstring = GetText();

            if (clipstring.Contains("-"))
            {
                InstalltionID = clipstring.Replace("-", "");
            }
            else if (clipstring.Contains(" "))
            {
                InstalltionID = clipstring.Replace(" ", "");
            }
            else
            {
                InstalltionID = clipstring;
            }
            InstalltionID = System.Text.RegularExpressions.Regex.Replace(InstalltionID, "[^\\d]", "");
            if (InstalltionID.Trim().Length == 54)
            {
                textBox3.Text = InstalltionID;
                button1.Enabled = true;
            }
            else if (InstalltionID.Trim().Length == 63)
            {
                textBox3.Text = InstalltionID;
                button1.Enabled = true;
            }
            else
            {
                textBox3.Enabled = false;
                
            }
        }

        public string GetText()
        {
            string ReturnValue = string.Empty;
            Thread STAThread = new Thread(() =>
            {
                ReturnValue = System.Windows.Forms.Clipboard.GetText();
            });
            STAThread.SetApartmentState(ApartmentState.STA);
            STAThread.Start();
            STAThread.Join();
            return ReturnValue;
        }
    }
}

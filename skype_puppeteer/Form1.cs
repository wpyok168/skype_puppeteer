using PuppeteerSharp.Input;
using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Threading;
using skype_puppeteer.Properties;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;

namespace skype_puppeteer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string name = "*****@gmail.com";
        private string pwd = "*********";
        private void Form1_Load(object sender, EventArgs e)
        {
            SkypeLogin();
        }
        private IBrowser browser;
        private IPage page;

        public async void SkypeLogin()
        {
            try
            {
                int width = Screen.PrimaryScreen.Bounds.Width;
                int height = Screen.PrimaryScreen.Bounds.Height;
                //await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultChromiumRevision); //自动下载chrome
                LaunchOptions option = new LaunchOptions();
                option.Headless = false;
                // "--start-maximized",
                option.Args = new string[] { "--use-fake-ui-for-media-stream", "--proxy-server=socks5://127.0.0.1:10808" };//浏览器窗口最大化 及使用代理服务器
                //option.DefaultViewport = new ViewPortOptions() { Width = width, Height = height };
                //string path = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command", null, null) as string;
                string path = "D:\\Program Files (x86)\\Chrome\\chrome.exe";
                if (path != null)
                {
                    option.ExecutablePath = path; //指定chrome.exe路径
                }
                //option.ExecutablePath = @"D:\Program Files\aa\1编程\c#\20190320\puppeteer\puppeteer\bin\Debug\.local-chromium\Win64-706915\chrome-win\chrome.exe";
                //option.WebSocketFactory = (uri, socketOptions, cancellationToken) =>
                //        System.Net.WebSockets.SystemClientWebSocket.ConnectAsync(uri, cancellationToken);
                //option.SlowMo = 100;


                browser = await Puppeteer.LaunchAsync(option);

                //await page.SetRequestInterceptionAsync(true);
                page = await browser.NewPageAsync();
                await page.GoToAsync("https://web.skype.com/");

                if (page.Url.Contains("login.live.com"))
                {
                    string url = System.Net.WebUtility.UrlDecode(page.Url);
                    if (!System.Net.WebUtility.UrlDecode(url).Contains("https://web.skype.com/Auth/PostHandle"))
                    {
                        var i0116 = await page.QuerySelectorAsync("#i0116");
                        //await page.EvaluateExpressionAsync(Resources.String1);
                        await page.WaitForSelectorAsync("#i0116");
                        await page.EvaluateExpressionAsync("document.getElementById(\"i0116\").setAttribute(\"value\",\"\")");
                        await page.TypeAsync("#i0116", name, new TypeOptions { Delay = 100 });
                        await page.ClickAsync("#idSIButton9", new ClickOptions { Delay = 1000 });

                        await page.WaitForTimeoutAsync(2000);
                        await page.WaitForSelectorAsync("#i0118");
                        await page.EvaluateExpressionAsync("document.getElementById(\"i0118\").setAttribute(\"value\",\"\")");
                        await page.EvaluateExpressionAsync("document.getElementById(\"i0118\").value=\"\"");
                        await page.TypeAsync("#i0118", pwd, new TypeOptions { Delay = 100 });
                        await page.WaitForTimeoutAsync(2000);
                        await page.ClickAsync("#idSIButton9", new ClickOptions { Delay = 1000 });

                        await page.WaitForNavigationAsync();
                        //await page.WaitForTimeoutAsync(2000);
                        await page.ClickAsync("#idSIButton9", new ClickOptions { Delay = 1000 });
                    }

                    
                }

                page.DOMContentLoaded += Page_DOMContentLoaded;
                page.Response += Page_Response;
                page.RequestFinished += Page_RequestFinished;
                await page.WaitForTimeoutAsync(10000);
                //await page.CloseAsync();
                //System.Threading.Thread.Sleep(50000);

                //document.querySelector('.css-1dbjc4n.r-18u37iz.r-1jkjb>div>button').click();
                //document.querySelector('[placeholder="电话号码"]').value="448000188354"
                //document.querySelector('[title="呼叫。"]').click();
                //document.querySelector('[title="继续但不启用音频或视频"]').click();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                
            }
        }

        private void Page_RequestFinished(object sender, RequestEventArgs e)
        {
            string url1 = e.Request.Url;
            //Debug.WriteLine(url1);
        }

        private async void Page_Response(object sender, ResponseCreatedEventArgs e)
        {
            string url = e.Response.Url;
            //Debug.WriteLine(url);

            if (e.Response.Status== HttpStatusCode.OK)
            {
                //azeus1-client-s.gateway.messenger.live.com
                if (e.Response.Url.Contains("-client-s.gateway.messenger.live.com"))
                {
                    //https://azwus1-client-s.gateway.messenger.live.com/v1/users/ME/endpoints/%7Bcdd6f119-8ac0-45ad-812c-8685c1e70851%7D/subscriptions/0/poll?ackId=1016
                    
                    try
                    {
                        string html = await e.Response.TextAsync();
                        JObject json = await e.Response.JsonAsync();
                       // Debug.Print(html);
                    }
                    catch (Exception)
                    {

                    }
                }
                
                

            }
            
        }

        private async void Page_DOMContentLoaded(object sender, EventArgs e)
        {
            if (page.Url.Equals("https://web.skype.com/"))
            {
                var element = await page.WaitForSelectorAsync(".css-1dbjc4n.r-18u37iz.r-1jkjb>div>button");
                await page.WaitForTimeoutAsync(2000);
                await page.EvaluateExpressionAsync("document.querySelector('.css-1dbjc4n.r-18u37iz.r-1jkjb>div>button').click()");
                //await page.ClickAsync(".css-1dbjc4n.r-18u37iz.r-1jkjb>div>button");
                await page.WaitForTimeoutAsync(2000);
                var result = await page.EvaluateFunctionAsync("()=>{var interval7 = setInterval(function() {\r\n        document.querySelectorAll(\"button\").forEach(ev=>{if(ev.getAttribute(\"title\")==\"继续\"){\r\n            //alert(ev.getAttribute(\"title\"));\r\n            ev.click();\r\n        }});},\r\n   1000);}");

                await page.WaitForTimeoutAsync(2000);
                //await page.WaitForSelectorAsync("[placeholder=\"电话号码\"]");
                await page.TypeAsync("[placeholder=\"电话号码\"]", "+448000188354", new TypeOptions { Delay = 100 });
                await page.WaitForTimeoutAsync(5000);
                await page.ClickAsync("[title=\"呼叫。\"]", new ClickOptions { Delay = 1000 });
                //
                await page.WaitForTimeoutAsync(2000);
                string noaudio = "noaudioclick();function noaudioclick() {\r\n    var interval6 = setInterval(function() {\r\n        document.querySelectorAll(\"button\").forEach(ev=>{if(ev.getAttribute(\"title\")==\"继续但不启用音频或视频\"){\r\n            //alert(ev.getAttribute(\"title\"));\r\n            ev.click();\r\n        }});},\r\n   1000);\r\n};";
                await page.EvaluateExpressionAsync(noaudio);
                string fy = "fy();\r\nfunction fy() {\r\n    var interval1 = setInterval(function() {\r\n        var fy = document.querySelector(\"div[data-text-as-pseudo-element='翻译此调用']\");\r\n        if (fy != null) {\r\n            clearInterval(interval1);\r\n            fy.click();\r\n        }\r\n    },\r\n    800);\r\n};";
                await page.EvaluateExpressionAsync(fy);
                string fyset = "fyset();\r\nfunction fyset() {\r\n    var interval1 = setInterval(function() {\r\n        var fy = document.querySelector(\"button[aria-label='翻译设置']\");\r\n        if (fy != null) {\r\n            clearInterval(interval1);\r\n            document.querySelector(\"button[aria-label='翻译设置']\").click();\r\n            document.querySelectorAll(\".css-1dbjc4n.r-1yf6mjq.r-1867qdf.r-1yadl64.r-13awgt0.r-18u37iz.r-xd6kpl.r-f727ji.r-j2kj52.r-tskmnb\")[0].click();\r\n            document.querySelector(\"button[aria-label='英语(英国)']\").click();\r\n            document.querySelectorAll(\".css-1dbjc4n.r-1yf6mjq.r-1867qdf.r-1yadl64.r-13awgt0.r-18u37iz.r-xd6kpl.r-f727ji.r-j2kj52.r-tskmnb\")[1].click();\r\n            document.querySelector(\"button[aria-label='英语(英国)']\").click();\r\n            document.querySelector(\"button[aria-label='完成']\").click();\r\n\r\n        }\r\n    },\r\n    800);\r\n};";
                await page.EvaluateExpressionAsync(fyset);

                while (true)
                {
                   var t = Task.Run(() => {
                       try
                       {
                           string jsstr = "var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\n    if(tq.length>0){\r\n        var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n        if(content.length>0){\r\n            var fy = content[content.length-1];\r\n            //alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n           fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n        }\r\n   }";
                           var ret = page.EvaluateFunctionAsync("()=>{var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\n    if(tq.length>0){\r\n        var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n        if(content.length>0){\r\n            var fy = content[content.length-1];\r\n            //alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n       return    fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n        }\r\n   }}");
                           
                           if (ret.Result!=null)
                           {
                               Debug.Print(ret.Result.ToString());
                           }
                          
                       }
                       catch (Exception)
                       {
                            
                       }
                        
                    });
                   t.Wait();
                    
                    System.Threading.Thread.Sleep(2000);

                }

                //System.Threading.Timer timer = new System.Threading.Timer(GetFY,null,500,2000);

                //System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                //timer.Interval = 1000;
                //timer.Enabled = true;
                //timer.Tick += Timer_Tick;
                //timer.Start();
                
                //string jsstr = "var interval1 = setInterval(function() {\r\n        var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\nif(tq.length>0){\r\n    var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n    if(content.length>0){\r\n        var fy = content[content.length-1];\r\n        alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n        return fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n    }\r\n}\r\n    },\r\n    800);";
               // var ret = page.EvaluateExpressionAsync(jsstr);
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                //string jsstr = "var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\nif(tq.length>0){\r\n    alert(\"a\");\r\n    var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n    if(content.length>0){\r\n        var fy = content[content.length-1];\r\n        alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n        return fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n    }\r\n}";
                string jsstr = "var interval1 = setInterval(function() {\r\n        var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\nif(tq.length>0){\r\n    alert(\"a\");\r\n    var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n    if(content.length>0){\r\n        var fy = content[content.length-1];\r\n        alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n        return fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n    }\r\n}\r\n    },\r\n    800);";
                var ret = page.EvaluateExpressionAsync(jsstr);
                Debug.WriteLine(ret.Result.ToString());
            }
            catch (Exception)
            {

            }
        }

        private void GetFY(object obj)
        {
            try
            {
                //string jsstr = "var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\nif(tq.length>0){\r\n    alert(\"a\");\r\n    var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n    if(content.length>0){\r\n        var fy = content[content.length-1];\r\n        alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n        return fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n    }\r\n}";
                string jsstr = "var tq =document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh\");\r\n    if(tq.length>0){\r\n        var content=document.querySelectorAll(\".css-1dbjc4n.r-150rngu.r-eqz5dr.r-16y2uox.r-1wbh5a2.r-11yh6sk.r-1rnoaur.r-2eszeu.r-1sncvnh>div>div\");\r\n        if(content.length>0){\r\n            var fy = content[content.length-1];\r\n            //alert(fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\"));\r\n            return fy.children[0].children[0].getAttribute(\"data-text-as-pseudo-element\");\r\n        }\r\n   }";
                var ret = page.EvaluateExpressionAsync(jsstr);
                Debug.WriteLine(ret.Result.ToString());
            }
            catch (Exception)
            {

            }
            
        }
    }
}

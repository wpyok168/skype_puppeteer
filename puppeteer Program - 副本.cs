using Newtonsoft.Json.Linq;
using PuppeteerSharp;
using PuppeteerSharp.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace puppeteer
{
    class Program
    {
        static void Main(string[] args)
        {
            JDTest().Wait();
            //Test().Wait();
            Console.ReadKey();
        }
        static async Task Test()
        {
            try
            {
                //await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision); //自动下载chrome
                LaunchOptions option = new LaunchOptions();
                option.Headless = false;
                //string path = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command", null, null) as string;
                //if (path != null)
                //{
                //    option.ExecutablePath = path; //指定chrome.exe路径
                //}
                //option.ExecutablePath = @"D:\Program Files\aa\1编程\c#\20190320\puppeteer\puppeteer\bin\Debug\.local-chromium\Win64-706915\chrome-win\chrome.exe";
                option.WebSocketFactory = (uri, socketOptions, cancellationToken) =>
                        System.Net.WebSockets.SystemClientWebSocket.ConnectAsync(uri, cancellationToken);
                option.SlowMo = 100;
                var browser = await Puppeteer.LaunchAsync(option);
                using (var page = await browser.NewPageAsync())
                {
                    await page.GoToAsync("https://www.baidu.com/");
                    var html = await page.EvaluateFunctionAsync<string>("() => document.documentElement.outerHTML");
                    var html1 = await page.GetContentAsync();

                    //await page.TypeAsync("#kw", "Puppeteer");
                    var t = await page.EvaluateExpressionAsync("document.getElementById(\"su\").outerHTML");//javascript脚本
                    string outerHTML = t.ToString();
                    var t1 = await page.QuerySelectorAsync("#kw");
                    var outerHTML1 = await t1.GetPropertyAsync("outerHTML");
                    var str = await outerHTML1.JsonValueAsync();

                    var outerHTML2 = await t1.GetPropertyAsync("id");
                    var id = await outerHTML2.JsonValueAsync();

                    var outerHTML3 = await t1.EvaluateFunctionAsync<string>("element => element.outerHTML");
                    var id1 = await t1.EvaluateFunctionAsync<string>("element => element.getAttribute('id')");

                    //page.WaitForXPathAsync
                    await page.TypeAsync("#kw", "Puppeteer");
                    await page.ClickAsync("#su",new ClickOptions() { Delay=1000}); //Delay 延迟输入
                    //string javascript = "document.getElementById(\"su\").click();";  
                    //await page.EvaluateExpressionAsync(javascript);//javascript脚本
                    await page.WaitForTimeoutAsync(2000);

                    //var tags = await page.EvaluateExpressionAsync<List<string>>("document.getElementsByTagName('a')");
                    var tags = await page.EvaluateFunctionAsync<List<string>>("() => {var tags = document.getElementsByTagName('a');var myas = new Array();for (var v = 0; v < tags.length; v++) {myas[v]=tags[v].outerHTML; } return myas;}");

                    var selector = await page.QuerySelectorAllHandleAsync("a");
                    var innerhtml = await selector.EvaluateFunctionAsync<string[]>("elements => elements.map(a=>a.innerHTML)");
                    var innerhtml1 = await selector.EvaluateFunctionAsync<string[]>("elements => elements.map(a=>a.outerHTML)");
                    var hanld = await selector.GetPropertiesAsync();
                    foreach (var item in hanld)
                    {
                        var ahtml = await item.Value.EvaluateFunctionAsync("elemnt=>elemnt.outerHTML");
                    }
                    var zeroinnerhtml = await hanld["0"].EvaluateFunctionAsync("elemnt=>elemnt.innerHTML");
                    var zeroinnerhtml1 = await hanld["0"].EvaluateFunctionAsync("elemnt=>elemnt.outerHTML");

                    await page.WaitForTimeoutAsync(10000);
                    await page.CloseAsync();
                    //System.Threading.Thread.Sleep(50000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        static async Task JDTest()
        {
            try
            {
                int width = Screen.PrimaryScreen.Bounds.Width;
                int height = Screen.PrimaryScreen.Bounds.Height;
                //await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision); //自动下载chrome
                LaunchOptions option = new LaunchOptions();
                option.Headless = false;
                option.Args = new string[] { "--start-maximized" };//浏览器窗口最大化
                option.DefaultViewport = new ViewPortOptions() { Width = width, Height = height };
                //string path = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command", null, null) as string;
                //if (path != null)
                //{
                //    option.ExecutablePath = path; //指定chrome.exe路径
                //}
                //option.ExecutablePath = @"D:\Program Files\aa\1编程\c#\20190320\puppeteer\puppeteer\bin\Debug\.local-chromium\Win64-706915\chrome-win\chrome.exe";
                option.WebSocketFactory = (uri, socketOptions, cancellationToken) =>
                        System.Net.WebSockets.SystemClientWebSocket.ConnectAsync(uri, cancellationToken);

                var browser = await Puppeteer.LaunchAsync(option);
                
                using (var page = await browser.NewPageAsync())
                {
                    await page.GoToAsync("https://passport.jd.com/new/login.aspx?ReturnUrl=https%3A%2F%2Fwww.jd.com%2F");
                    await page.WaitForSelectorAsync("#loginname");
                    await page.ClickAsync("#content > div.login-wrap > div.w > div > div.login-tab.login-tab-r > a");
                    await page.TypeAsync("#loginname", "账号",new TypeOptions() { Delay=100});
                    await page.TypeAsync("#nloginpwd", "密码", new TypeOptions() { Delay = 100 });
                    await page.ClickAsync("#loginsubmit");
                    await page.WaitForSelectorAsync("#JDJRV-wrap-loginsubmit > div");
                    //while (true)
                    //{
                    //    try
                    //    {
                    //        await page.WaitForXPathAsync("//*[@id='JDJRV-wrap-loginsubmit']/div/div/div/div[1]/div[2]/div[1]/img");
                    //        await page.WaitForTimeoutAsync(3000);
                    //        var DJRV_bigimg = await page.WaitForSelectorAsync("#JDJRV-wrap-loginsubmit > div > div > div > div.JDJRV-img-panel.JDJRV-click-bind-suspend > div.JDJRV-img-wrap > div.JDJRV-bigimg > img");
                    //        //var DJRV_bigimg = await page.QuerySelectorAsync("#JDJRV-wrap-loginsubmit > div > div > div > div.JDJRV-img-panel.JDJRV-click-bind-suspend > div.JDJRV-img-wrap > div.JDJRV-bigimg > img");
                    //        string imgbase64str = await DJRV_bigimg.EvaluateFunctionAsync<string>("element => element.getAttribute('src')");
                    //        if (string.IsNullOrEmpty(imgbase64str))
                    //        {
                    //            break;
                    //        }
                    //        var JDJRV_smallimg = await page.WaitForSelectorAsync("#JDJRV-wrap-loginsubmit > div > div > div > div.JDJRV-img-panel.JDJRV-click-bind-suspend > div.JDJRV-img-wrap > div.JDJRV-smallimg > img");
                    //        BoundingBox yid = await JDJRV_smallimg.BoundingBoxAsync();

                    //        //var DJRV_bigimg = await page.QuerySelectorAsync("#JDJRV-wrap-loginsubmit > div > div > div > div.JDJRV-img-panel.JDJRV-click-bind-suspend > div.JDJRV-img-wrap > div.JDJRV-bigimg > img");
                    //        string imgbase64str1 = await JDJRV_smallimg.EvaluateFunctionAsync<string>("element => element.getAttribute('src')");

                    //        Image64ToImage(imgbase64str, "DJRV_bigimg");
                    //        Image64ToImage(imgbase64str1, "JDJRV_smallimg");
                    //        //======================

                    //        Mat mat1 = new Mat(System.Environment.CurrentDirectory + "\\DJRV_bigimg.png", ImreadModes.AnyDepth);
                    //        //Cv2.ImShow("mat1", mat1);
                    //        Mat mat2 = new Mat(System.Environment.CurrentDirectory + "\\JDJRV_smallimg.png", ImreadModes.AnyDepth);
                    //        Mat mat3 = new Mat();
                    //        //mat3.Create(mat1.Cols - mat2.Cols + 1, mat1.Rows - mat2.Cols + 1, MatType.CV_32FC1);
                    //        Cv2.MatchTemplate(mat1, mat2, mat3, TemplateMatchModes.CCoeffNormed);
                    //        //Cv2.Normalize(mat3, mat3, 1, 0, NormTypes.MinMax, -1);
                    //        OpenCvSharp.Point minloc, maxloc;
                    //        Cv2.MinMaxLoc(mat3, out minloc, out maxloc);
                    //        int w1 = minloc.X;

                    //        try
                    //        {
                    //            await page.Mouse.MoveAsync(yid.X + yid.Width / 2, yid.Y + yid.Height / 2);
                    //        }
                    //        catch (Exception)
                    //        {
                    //            break;
                    //        }
                    //        await page.Mouse.DownAsync();
                    //        decimal currentpoit = yid.X + w1;
                    //        await page.Mouse.MoveAsync(currentpoit, yid.Y + yid.Height / 2 + 10, new MoveOptions() { Steps = 80 });
                    //        await page.WaitForTimeoutAsync(1000);
                    //        currentpoit = currentpoit + 3;
                    //        await page.Mouse.MoveAsync(currentpoit, yid.Y + yid.Height / 2 - 6, new MoveOptions() { Steps = 120 });
                    //        await page.WaitForTimeoutAsync(1000);
                    //        currentpoit = currentpoit - 1;
                    //        await page.Mouse.MoveAsync(currentpoit, yid.Y + yid.Height / 2 - 10, new MoveOptions() { Steps = 160 });
                    //        await page.WaitForTimeoutAsync(1000);
                    //        await page.Mouse.UpAsync();
                    //        //await page.Mouse.ClickAsync(yid.X + w1 - 25,yid.Y + yid.Height / 2);
                    //        //await page.Mouse.ClickAsync(0, 0);
                    //    }
                    //    catch (Exception)
                    //    {
                    //        break;
                    //    }
                    //}


                    //==============
                    await page.WaitForSelectorAsync("#key");
                    //while (true)
                    //{
                    //    if (page.Url.Equals("https://www.jd.com/"))
                    //    {
                    //        break;
                    //    }
                    //}
                    //await page.TypeAsync("#key", "硬盘", new TypeOptions() { Delay = 100 });
                    //await page.ClickAsync("#search > div > div.form > button");

                    await page.GoToAsync("https://a.jd.com/");
                    await page.EvaluateExpressionAsync("window.scrollTo(0,1000)");//页面滚动
                    try
                    {
                        await page.WaitForSelectorAsync("#quanlist > div.quan-item.quan-item02.quan-type01.item-710721250.quan-item-unload > div.q-ops-box > div > a");
                        for (int i = 0; i < 10; i++)
                        {
                            await page.ClickAsync("#quanlist > div.quan-item.quan-item02.quan-type01.item-710721250.quan-item-unload > div.q-ops-box > div > a", new ClickOptions() { Delay = 200 });
                            //await page.WaitForTimeoutAsync(100);
                            ElementHandle elementHandle = await page.QuerySelectorAsync("#quan-dialog > div > div > div.tip-btnbox > a");
                            if (elementHandle != null)
                            {
                                await elementHandle.ClickAsync();
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }


                    //await page.ClickAsync("#settleup > div", new ClickOptions() { Delay = 100 });

                    ////跳转新页面方法
                    ////await page.GoToAsync("https://cart.jd.com/cart.action");  //直接跳转新页面
                    ////await page.ClickAsync("#J_zypromo_btn");//直接跳转新页面

                    ////Target target = await browser.WaitForTargetAsync(t => t.Url == "https://cart.jd.com/cart.action"); //通过browser.waitForTarget获取target 
                    ////var page1 = await target.PageAsync();//通过browser.waitForTarget获取target
                    ////await page1.WaitForSelectorAsync("#J_zypromo_btn");//通过browser.waitForTarget获取target
                    ////await page1.ClickAsync("#J_zypromo_btn");//通过browser.waitForTarget获取target

                    ////browser.pages()可以获取所有打开的Page对象，可以通过遍历或筛选找到自己想获取的Page对象
                    //await page.WaitForTimeoutAsync(3000);//等待页面加载
                    //Page[] pages = await browser.PagesAsync(); //获取所有页面
                    //var page2 = pages.FirstOrDefault(t => t.Target.Url == "https://cart.jd.com/cart.action");
                    //await page2.WaitForSelectorAsync("#J_zypromo_btn");
                    //await page2.ClickAsync("#J_zypromo_btn");
                    ////browser.pages()可以获取所有打开的Page对象，可以通过遍历或筛选找到自己想获取的Page对象

                    await page.WaitForTimeoutAsync(50000);
                    await page.CloseAsync();
                    //System.Threading.Thread.Sleep(50000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        static Image Image64ToImage(string base64imagestr,string filename)
        {
            base64imagestr = base64imagestr.Replace("data:image/png;base64,", "");
            byte[] bytes = Convert.FromBase64String(base64imagestr);
            Image image;
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                image = Image.FromStream(ms);
                string path = System.Environment.CurrentDirectory + $"\\{filename}.png";
                image.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
            return image;
        }
    }
}

using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WindowsInput;

class MainClass
{
    static IWebDriver driver = null;
    static string currentEnvironmentPath = Path.GetFullPath(Path.Combine(@$"{Environment.CurrentDirectory}", @"..\..\..\"));
    static string taskFolder = Path.GetFullPath(Path.Combine(@$"{Environment.CurrentDirectory}", @"..\..\..\Misc"));
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    public struct POINT
    {
        public int X;
        public int Y;
    }
    public static void Main(string[] args)
    {
        #region Init variables
        SheetsService _googleSheetService = ImportantMethods.InitializeGoogleSheet();
        IList<IList<object>> sheetData = null;
        var errorMessages = new List<string>();
        var oneLoopDone = false;
        driver = ImportantMethods.WebDriver(false, false, "chrome", true);
        #endregion
        #region Monitoring Algorithm
        while (true)
        {
            while (true) { try { sheetData = ImportantMethods.ReadSpreadsheetEntries($"Orders!A1:X10000", _googleSheetService, "1nwe4Xcc6tcaa3MA79_yls9B9dmQ2hvITcguaaEciX4c"); break; } catch (Exception) { Thread.Sleep(TimeSpan.FromMinutes(1)); } }
            for (int i = 0; i<sheetData.Count; i++)
            {
                try
                {
                    var currentOrderObj = ReturnOrderObj(sheetData, i);
                    if (currentOrderObj.OrderStatus == "created")
                    {
                        //search for the product via the "SamsClubSku" value
                        //ImportantMethods.VisitSite(driver, "https://www.samsclub.com/?xid=hdr_logo");
                        if (oneLoopDone == true)
                        {
                            try { driver.FindElement(By.CssSelector("a[class=\"sc-simple-header-icon\"]")).Click(); } catch (Exception) { driver.FindElement(By.CssSelector("a[class=\"logo\"]")).Click(); }
                        }
                        //driver.FindElement(By.CssSelector("input[type=\"search\"]")).SendKeys(Keys.Control + "a");
                        //driver.FindElement(By.CssSelector("input[type=\"search\"]")).SendKeys(currentOrderObj.SamsClubSku);
                        //Thread.Sleep(1000);
                        //driver.FindElement(By.CssSelector("button[class=\"sc-search-field-icon\"]")).Click();
                        ImportantMethods.VisitSite(driver, $"https://www.samsclub.com/s/{currentOrderObj.SamsClubSku}");
                        Thread.Sleep(5000);
                        //getting product count and notify if 0 products have been found or more than 1
                        var productCount = driver.FindElements(By.CssSelector("a[role=\"group\"]")).Count;
                        //if cart contains > 1 products, notify via email and wait for further instructions
                        if (driver.FindElements(By.CssSelector("div[class=\"cart-label-container\"]")).Count()>0)
                        {
                            ClickOnCart();
                            Thread.Sleep(5000);
                            var allRemoveBtn = driver.FindElements(By.CssSelector("div[class=\"sc-cart-item-actions sc-cart-item-actions-border\"] > button:nth-child(1)"));
                            foreach (var btn in allRemoveBtn)
                            {
                                btn.Click();
                                Thread.Sleep(5000);
                            }
                            //ImportantMethods.SendEmail($"[Notice for JSF]: Cart cleared out because there was a item before automating", $"Hi, <br> product(s) are already in the cart, I removed that in order to continue.", "wilson@jsfproducts.com");
                            ImportantMethods.VisitSite(driver, $"https://www.samsclub.com/s/{currentOrderObj.SamsClubSku}");
                            productCount = driver.FindElements(By.CssSelector("a[role=\"group\"]")).Count;
                        }
                        //if more than one product appears then wait until further notice
                        if (productCount > 1)
                        {
                            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastErrorScreenshot.png", ScreenshotImageFormat.Png);
                            ImportantMethods.SendEmail($"[Action Required for JSF]: Multiple products found on search result page for: '{currentOrderObj.SamsClubSku}'", $"Hi, <br> multiple products appeared on this search result page, what shall the bot do? Let Suheyl know so he can make the application continue. <br> Link: https://www.samsclub.com/s/{currentOrderObj.SamsClubSku}", "wilson@jsfproducts.com", "lastErrorScreenshot.png");
                        }
                        //if product count is 0 notify via email so further action can be taken
                        else if (productCount == 0)
                        {
                            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastErrorScreenshot.png", ScreenshotImageFormat.Png);
                            ImportantMethods.SendEmail($"[Action Required for JSF]: No product found for: '{currentOrderObj.SamsClubSku}'", $"Hi, <br> the bot did not find any products for the search result page: https://www.samsclub.com/s/{currentOrderObj.SamsClubSku} <br> Let Suheyl know (He may also have read this email) and further action can be taken.", "wilson@jsfproducts.com", "lastErrorScreenshot.png");
                            continue;
                        }
                        //add the first product appearing to cart from the search result page
                        Thread.Sleep(5000);
                        if (driver.FindElements(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block sc-pc-action-button sc-pc-out-of-stock-button\"]")).Count>0)
                        {
                            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastErrorScreenshot.png", ScreenshotImageFormat.Png);
                            ImportantMethods.SendEmail($"[Action Required for JSF]: This product is out of stock: '{currentOrderObj.SamsClubSku}'", $"Hi, <br> the bot searched for a product which is out of stock, here's the search result page: https://www.samsclub.com/s/{currentOrderObj.SamsClubSku} <br> Let Suheyl know (He may also have read this email) and further action can be taken.", "wilson@jsfproducts.com", "lastErrorScreenshot.png");
                            continue;
                        }
                       driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block sc-pc-action-button sc-pc-add-to-cart\"]")).Click();
                        //visit checkout page
                        Thread.Sleep(5000);
                        bool yesSkuInCartIsCorrect = false;
                        while (true)
                        {
                            ClickOnCart();
                            //checking if items on cart are the same as the item that should be bought at the moment
                            var allSkus = driver.FindElements(By.CssSelector("p[class=\"sc-cart-item-number\"]")).Select(x => x.GetAttribute("innerText")).ToList();
                            for (int b = 0; b < allSkus.Count; b++)
                            {
                                if (allSkus[b].Contains(currentOrderObj.SamsClubSku))
                                {
                                    yesSkuInCartIsCorrect =true;
                                }
                                break;
                            }
                            if (yesSkuInCartIsCorrect)
                            {
                                //clicking on "Begin checkout" button on the left panel
                                Thread.Sleep(5000);
                                driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block sc-cart-begin-checkout-button\"]")).Click();
                                //clicking on "change adress button" to change the adress to the customers adress
                                Thread.Sleep(12000);
                                driver.FindElement(By.CssSelector("button[class=\"sc-btn fake-link sc-order-address-change\"]")).Click();
                                //logic to use existing shipping adress if present or create it if necessary
                                {
                                    //create address
                                    //clicking on add adress btn
                                    driver.FindElement(By.CssSelector("button[class=\"sc-btn fake-link sc-plus-button\"]")).Click();
                                    //entering the customers adress
                                    //first & last name
                                    driver.FindElement(By.CssSelector("input[aria-label=\"First and last name\"]")).SendKeys(currentOrderObj.name);
                                    //entering adress1 into the "enter street" input box
                                    driver.FindElement(By.CssSelector("input[name=\"addressLineOne\"]")).SendKeys($"{currentOrderObj.address1}");
                                    //entering adress2 into the "enter street" input box
                                    driver.FindElement(By.CssSelector("input[name=\"addressLineTwo\"]")).SendKeys($"{currentOrderObj.address2}");
                                    //entering city
                                    driver.FindElement(By.CssSelector("div[class=\"sc-input-box sc-address-fields-city\"] > div > input")).SendKeys($"{currentOrderObj.city}");
                                    //selecting state
                                    new SelectElement(driver.FindElement(By.CssSelector("select[class=\"visuallyhidden\"]"))).SelectByValue(currentOrderObj.state);
                                    //entering postal code
                                    driver.FindElement(By.CssSelector("div[class=\"sc-address-fields-zip\"] > div > div > input")).SendKeys($"{currentOrderObj.postalCode}");
                                    //entering phone number
                                    driver.FindElement(By.CssSelector("div[class=\"sc-masked-input-box\"] > div > div > input")).SendKeys($"{currentOrderObj.phone}");
                                    //clicking save btn
                                    driver.FindElement(By.CssSelector("button[type=\"submit\"]")).Click();
                                    CheckIfTooManyAddresses();
                                    //check if "sign in pop up is present"
                                    var needToClickOnSaveBtn = true;
                                    try
                                    {
                                        Thread.Sleep(2000);
                                        //clicking sign in btn
                                        driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block\"]")).Click();
                                        Thread.Sleep(5000);
                                        //checking if it want to confirm via code
                                        if (driver.PageSource.Contains("For your security, we need to confirm it's you"))
                                        {
                                            driver.FindElement(By.CssSelector("body > div:nth-child(8) > div > div > div > div.sc-modal-content-scrollable > div > div > div:nth-child(1) > div > div > div:nth-child(1) > div > div > div.sc-2fa-enroll-container > form > ul > li:nth-child(2) > div > label")).Click();
                                            driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block\"]")).Click();
                                            ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                                            driver.Navigate().GoToUrl("https://outlook.office365.com/mail/");
                                            Thread.Sleep(35000);
                                            var allEmails = driver.FindElements(By.CssSelector("div[data-is-scrollable=\"true\"] > div > div[aria-selected=\"false\"]"));
                                            var code = "";
                                            foreach (var email in allEmails)
                                            {
                                                if (email.GetAttribute("innerText").Contains("Sam's Club"))
                                                {
                                                    email.Click();
                                                    break;
                                                }
                                            }
                                            var tempList = driver.FindElements(By.CssSelector("td[style=\"padding:0 0 18px\"] > p[style]"));
                                            foreach (var item in tempList)
                                            {
                                                var match = Regex.Match(item.GetAttribute("innerText"), @"\d\d\d\d\d\d");
                                                if (match.Success)
                                                {
                                                    code = match.Value;
                                                    break;
                                                }
                                            }
                                            ((IJavaScriptExecutor)driver).ExecuteScript("window.close();");
                                            driver.SwitchTo().Window(driver.WindowHandles.First());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[0].SendKeys(code[0].ToString());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[1].SendKeys(code[1].ToString());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[2].SendKeys(code[2].ToString());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[3].SendKeys(code[3].ToString());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[4].SendKeys(code[4].ToString());
                                            driver.FindElements(By.CssSelector("input[aria-label=\"Passcode digit\"]"))[5].SendKeys(code[5].ToString());
                                            driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-btn-block\"]")).Click();
                                            needToClickOnSaveBtn = false;
                                            Thread.Sleep(10000);
                                            continue;
                                        }
                                    }
                                    catch (Exception) { }
                                    //check if sams club says "address can't be verified"
                                    if (driver.FindElements(By.CssSelector("div[class=\"sc-alert sc-alert-update\"]")).Count > 0)
                                    {
                                        ImportantMethods.SendEmail($"[Action Required for JSF]: Sams Club says address can't be verified. Check this", $"Hi, <br> the bot entered a adress and after that Sams Club said it can't be verified, please check if everything was correct and continue the application", "wilson@jsfproducts.com");
                                        try { driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary\"]")).Click(); } catch (Exception) { }
                                        CheckIfTooManyAddresses();
                                    }
                                    else if (needToClickOnSaveBtn)
                                    {
                                        //click save btn
                                        try { driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary\"]")).Click(); } catch (Exception) { }
                                    }
                                }
                                break;
                            }
                            else
                            {
                                ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastErrorScreenshot.png", ScreenshotImageFormat.Png);
                                ImportantMethods.SendEmail($"[Action Required for JSF]: Unrecognized Item in cart.", $"Hi, <br> the cart contains a unrecognized item in the cart, please check the server and see what's up or contact Suheyl.", "wilson@jsfproducts.com", "lastErrorScreenshot.png");
                            }
                        }
                        //scraping deliveryDate
                        currentOrderObj.estimatedDeliveryDate = driver.FindElement(By.CssSelector("div[class=\"sc-checkout-group-box-message\"]")).GetAttribute("innerText");
                        //finally click the "Place order" btn
                        //insert pin if necessary/required
                        //try { driver.FindElement(By.CssSelector("input[type=\"password\"]")).SendKeys("2714"); } catch (Exception) { }
                        //driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-place-order\"]")).Click();
                        //reporting the order placement via email
                        {
                            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastScreenshot.png", ScreenshotImageFormat.Png);
                            ImportantMethods.SendEmail($"[No Action Required for JSF]: A order has successfully been placed.",
                                $"Hi, <br> a order with these parameters have been successfully placed: " +
                                $"<p><b>purchaseOrderID:</b> {currentOrderObj.purchaseOrderID}</p>"+
                                $"<p><b>customerOrderId:</b> {currentOrderObj.customerOrderId}</p>"+
                                $"<p><b>customerEmailId:</b> {currentOrderObj.customerEmailId}</p>"+
                                $"<p><b>orderDate:</b> {currentOrderObj.orderDate}</p>"+
                                $"<p><b>phone:</b> {currentOrderObj.phone}</p>"+
                                $"<p><b>estimatedDeliveryDate:</b> {currentOrderObj.estimatedDeliveryDate}</p>"+
                                $"<p><b>estimatedShipDate:</b> {currentOrderObj.estimatedShipDate}</p>"+
                                $"<p><b>methodCode:</b> {currentOrderObj.methodCode}</p>"+
                                $"<p><b>name:</b> {currentOrderObj.name}</p>"+
                                $"<p><b>address1:</b> {currentOrderObj.address1}</p>"+
                                $"<p><b>address2:</b> {currentOrderObj.address2}</p>"+
                                $"<p><b>city:</b> {currentOrderObj.city}</p>"+
                                $"<p><b>currency:</b> {currentOrderObj.currency}</p>"+
                                $"<p><b>state:</b> {currentOrderObj.state}</p>"+
                                $"<p><b>postalCode:</b> {currentOrderObj.postalCode}</p>"+
                                $"<p><b>country:</b> {currentOrderObj.country}</p>"+
                                $"<p><b>addressType:</b> {currentOrderObj.addressType}</p>"+
                                $"<p><b>chargeAmount:</b> {currentOrderObj.chargeAmount}</p>"+
                                $"<p><b>currency:</b> {currentOrderObj.currency}</p>"+
                                $"<p><b>taxAmount:</b> {currentOrderObj.taxAmount}</p>"+
                                $"<p><b>taxCurrency:</b> {currentOrderObj.taxCurrency}</p>"+
                                $"<p><b>productName:</b> {currentOrderObj.productName}</p>"+
                                $"<p><b>sku:</b> {currentOrderObj.sku}</p>"+
                                $"<p><b>SamsClubSku:</b> {currentOrderObj.SamsClubSku}</p>"+
                                $"<p><b>OrderStatus:</b> {currentOrderObj.OrderStatus}</p>",
                                "wilson@jsfproducts.com", "lastScreenshot.png");
                        }
                        ImportantMethods.UpdateSpreadsheet($"Orders!X{i + 1}:X{i + 1}", new List<object>() { "order placed" }, _googleSheetService, "1nwe4Xcc6tcaa3MA79_yls9B9dmQ2hvITcguaaEciX4c");
                        ImportantMethods.UpdateSpreadsheet($"Orders!F{i + 1}:F{i + 1}", new List<object>() { currentOrderObj.estimatedDeliveryDate }, _googleSheetService, "1nwe4Xcc6tcaa3MA79_yls9B9dmQ2hvITcguaaEciX4c");
                        oneLoopDone = true;
                    }
                    
                }
                catch (Exception e)
                {
                    var errorMsg = $"Error time: {DateTime.Now}, i = {i}/{sheetData.Count()}, Error msg: {e.ToString()}";

                    if (driver.PageSource.Contains("Let us know you’re human (no robots allowed)"))
                    {
                        SolveCaptcha();
                        i--;
                        continue;
                    }

                    ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile("lastErrorScreenshot.png", ScreenshotImageFormat.Png);
                    File.AppendAllLines($@"{taskFolder}\log.txt", new string[] { errorMsg });
                    ImportantMethods.SendEmail($"[Action Required for JSF]: A unhandled exception occurred.", $"Hi, <br> the bot is experiencing a unhandled exception right now, here are the details, Suheyl will know what to do: <br> {errorMsg}", "wilson@jsfproducts.com", "lastErrorScreenshot.png");

                    errorMessages.Add(errorMsg);
                }
            }
            Thread.Sleep(TimeSpan.FromSeconds(10));
        }
        #endregion
    }
    public static void CheckIfTooManyAddresses()
    {
        //check if it says that there are too many adresses
        var btnSelector = "button[class=\"sc-btn fake-link sc-expanding-list-box-with-offset-show\"]";
        if (driver.PageSource.Contains("To add an address, you need to remove one first. You can save up to 10 addresses"))
        {
            driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-secondary\"]")).Click();
            while (driver.FindElements(By.CssSelector(btnSelector)).Count>0)
            {
                driver.FindElement(By.CssSelector(btnSelector)).Click();
            }
            while (driver.FindElements(By.CssSelector("button[class=\"sc-btn fake-link sc-address-card-delete-action\"]")).Count()>0)
            {
                driver.FindElement(By.CssSelector("button[class=\"sc-btn fake-link sc-address-card-delete-action\"]")).Click();
                Thread.Sleep(500);
                driver.FindElement(By.CssSelector("button[class=\"sc-btn sc-btn-primary sc-are-you-sure-button-yes\"]")).Click();
                Thread.Sleep(1500);
            }
        }
    }
    public static void ClickOnCart()
    {
        //clicking on "Begin checkout" button on the left panel
        driver.FindElement(By.CssSelector("div[class=\"container-cart\"]")).Click();
    }
    public static void SolveCaptcha()
    {
        while (true)
        {
            var coords = JsonConvert.DeserializeObject<List<POINT>>(File.ReadAllText($@"{taskFolder}\coords.txt"));
            var color = GetColorAt(1079, 604);
            var x = 34.16856492027335;
            var y = 60.70826306913997;
            var simulator = new InputSimulator();
            foreach (POINT coord in coords)
            {
                SetCursorPos(coord.X, coord.Y);
                System.Threading.Thread.Sleep(1);
                if (Console.KeyAvailable) break;
            }
            simulator.Mouse.LeftButtonDown();
            var stopwatch2 = new Stopwatch();
            stopwatch2.Start();
            if (driver.PageSource.Contains("Let us know you’re human (no robots allowed)"))
            {

                while (color.Name == "ffffffff")
                {
                    color = GetColorAt(1079, 604);
                }
                Thread.Sleep(Convert.ToInt32(stopwatch2.Elapsed.TotalSeconds * 50));
                simulator.Mouse.LeftButtonUp();
                coords.Reverse();
                foreach (POINT coord in coords)
                {
                    SetCursorPos(coord.X, coord.Y);
                    System.Threading.Thread.Sleep(1);
                    if (Console.KeyAvailable) break;
                }
                Thread.Sleep(10000);
            }
            if (driver.PageSource.Contains("Let us know you’re human (no robots allowed)"))
            {
                //simulator.Mouse.MoveMouseTo(x*1117, y*474);
                //simulator.Mouse.LeftButtonClick();
                driver.Navigate().Refresh();
                //if (driver.PageSource.Contains("Let us know you’re human (no robots allowed)"))
                {
                    //ImportantMethods.SendEmail($"[Action Required for JSF]: Captcha could not automatically be solved, manual attention needed.", $"Hi, <br> the bot sadly could not solve the captcha automatically and therefore needs manual attention. Please connect to the server and try to solve it.", "wilson@jsfproducts.com");
                    //driver.Close();
                    //driver.Quit();
                    //foreach (var item in System.Diagnostics.Process.GetProcessesByName("chromedriver")) { item.Kill(); }
                    //new DirectoryInfo($@"{currentEnvironmentPath}\ChromeBinary").Delete(true);
                    //Copy($@"{currentEnvironmentPath}\ChromeBinary Backup", $@"{currentEnvironmentPath}\ChromeBinary");
                    //driver = ImportantMethods.WebDriver(false, false, "chrome", true);
                }
            }
            else
            {
                break;
            }
        }
    }
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetWindowDC(IntPtr window);
    [DllImport("gdi32.dll", SetLastError = true)]
    public static extern uint GetPixel(IntPtr dc, int x, int y);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int ReleaseDC(IntPtr window, IntPtr dc);

    public static Color GetColorAt(int x, int y)
    {
        IntPtr desk = GetDesktopWindow();
        IntPtr dc = GetWindowDC(desk);
        int a = (int)GetPixel(dc, x, y);
        ReleaseDC(desk, dc);
        return Color.FromArgb(255, (a >> 0) & 0xff, (a >> 8) & 0xff, (a >> 16) & 0xff);
    }
    public static void Copy(string sourceDirectory, string targetDirectory)
    {
        var diSource = new DirectoryInfo(sourceDirectory);
        var diTarget = new DirectoryInfo(targetDirectory);

        CopyAll(diSource, diTarget);
    }

    public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
    {
        Directory.CreateDirectory(target.FullName);

        // Copy each file into the new directory.
        foreach (FileInfo fi in source.GetFiles())
        {
            fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
        }

        // Copy each subdirectory using recursion.
        foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
        {
            DirectoryInfo nextTargetSubDir =
                target.CreateSubdirectory(diSourceSubDir.Name);
            CopyAll(diSourceSubDir, nextTargetSubDir);
        }
    }
    static OrderObj ReturnOrderObj(IList<IList<object>> sheetData, int i)
    {
        var orderObj = new OrderObj();
        try { orderObj.purchaseOrderID = sheetData[i][0].ToString(); } catch (Exception) { }
        try { orderObj.customerOrderId = sheetData[i][1].ToString(); } catch (Exception) { }
        try { orderObj.customerEmailId = sheetData[i][2].ToString(); } catch (Exception) { }
        try { orderObj.orderDate = sheetData[i][3].ToString(); } catch (Exception) { }
        try { orderObj.phone = sheetData[i][4].ToString(); } catch (Exception) { }
        try { orderObj.estimatedDeliveryDate = sheetData[i][5].ToString(); } catch (Exception) { }
        try { orderObj.estimatedShipDate = sheetData[i][6].ToString(); } catch (Exception) { }
        try { orderObj.methodCode = sheetData[i][7].ToString(); } catch (Exception) { }
        try { orderObj.name = sheetData[i][8].ToString(); } catch (Exception) { }
        try { orderObj.address1 = sheetData[i][9].ToString(); } catch (Exception) { }
        try { orderObj.address2 = sheetData[i][10].ToString(); } catch (Exception) { }
        try { orderObj.city = sheetData[i][11].ToString(); } catch (Exception) { }
        try { orderObj.state = sheetData[i][12].ToString(); } catch (Exception) { }
        try { orderObj.postalCode = sheetData[i][13].ToString(); } catch (Exception) { }
        try { orderObj.country = sheetData[i][14].ToString(); } catch (Exception) { }
        try { orderObj.addressType = sheetData[i][15].ToString(); } catch (Exception) { }
        try { orderObj.chargeAmount = sheetData[i][16].ToString(); } catch (Exception) { }
        try { orderObj.currency = sheetData[i][17].ToString(); } catch (Exception) { }
        try { orderObj.taxAmount = sheetData[i][18].ToString(); } catch (Exception) { }
        try { orderObj.taxCurrency = sheetData[i][19].ToString(); } catch (Exception) { }
        try { orderObj.productName = sheetData[i][20].ToString(); } catch (Exception) { }
        try { orderObj.sku = sheetData[i][21].ToString(); } catch (Exception) { }
        try { orderObj.SamsClubSku = sheetData[i][22].ToString(); } catch (Exception) { }
        try { orderObj.OrderStatus = sheetData[i][23].ToString(); } catch (Exception) { }
        return orderObj;
    }
}
class OrderObj
{
    public string purchaseOrderID { get; set; }
    public string customerOrderId { get; set; }
    public string customerEmailId { get; set; }
    public string orderDate { get; set; }
    public string phone { get; set; }
    public string estimatedDeliveryDate { get; set; }
    public string estimatedShipDate { get; set; }
    public string methodCode { get; set; }
    public string name { get; set; }
    public string address1 { get; set; }
    public string address2 { get; set; }
    public string city { get; set; }
    public string state { get; set; }
    public string postalCode { get; set; }
    public string country { get; set; }
    public string addressType { get; set; }
    public string chargeAmount { get; set; }
    public string currency { get; set; }
    public string taxAmount { get; set; }
    public string taxCurrency { get; set; }
    public string productName { get; set; }
    public string sku { get; set; }
    public string SamsClubSku { get; set; }
    public string OrderStatus { get; set; }
}
using System.Text;
using OfficeOpenXml;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
namespace TestingProgram;

public static class Testing
{
    const string charSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    const string charSetWithSpecialChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789,.()";
    public static void RunAddProductTest(int numTest)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        GenerateResultFile();
        IWebDriver? driver;
        var options = new EdgeOptions();
        driver = new EdgeDriver(options);
        Random random = new Random();
        driver.Navigate().GoToUrl("http://localhost:58473/Admin/Login");
        WebDriverWait waitLoginPage = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        driver.FindElement(By.Id("exampleInputEmail")).Click();
        driver.FindElement(By.Id("exampleInputEmail")).SendKeys("admin");
        driver.FindElement(By.Id("exampleInputPassword")).Click();
        driver.FindElement(By.Id("exampleInputPassword")).SendKeys("admin");
        driver.FindElement(By.CssSelector(".btn-primary")).Click();
        WebDriverWait waitLoginTime = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        string[] cateSelections =
        {
            "Sách-Truyện tranh", "Sách kinh tế", "Sách kĩ năng sống", "Sách ngoại ngữ", "Sách thiếu nhi", "Dụng cụ vẽ",
            "Dụng cụ học tập"
        };
        string[] providerSelections =
        {
            "Kim Đồng", "Nhà Xuất Bản Giáo Dục Việt Nam", "Fahasa", "Nhà Xuất Bản Trẻ", "Alpha Books",
            "Nhà Xuất Bản Văn Học", "Saigon Books", "Đinh Tị Books", "Nhà Sách Đông A"
        };
        string[] imgpath =
        {
            "F:\\Learn\\TMDT\\websitebandodunghoctap\\WebsiteDoDungHocTap\\TestingProgram\\Img_test\\1.png",
            "F:\\Learn\\TMDT\\websitebandodunghoctap\\WebsiteDoDungHocTap\\TestingProgram\\Img_test\\2.jpg",
            "F:\\Learn\\TMDT\\websitebandodunghoctap\\WebsiteDoDungHocTap\\TestingProgram\\Img_test\\3.mp4",
            ""
        };
        #region test name
        
        //empty name
        for (int i = 0; i < numTest; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                try
                {
                    var isValid =
                        (bool)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].checkValidity();",
                            nameElement);
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                catch (WebDriverException e)
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                    
                }
                worksheet.Cells[row, 13].Value = "Name";
                package.Save();
            #endregion
        }
        //legal random name
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name

                int length = random.Next(1, 50);
                string namestr = GenerateRandomSpecialString(length);
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys(namestr);
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = namestr;
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Name";
                package.Save();
            #endregion
        }
        //>50 length
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name

                int length = random.Next(50, 75);
                string namestr = GenerateRandomSpecialString(length);
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys(namestr);
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = namestr;
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                IWebElement validationErrorElement = driver.FindElement(By.CssSelector("span[data-valmsg-for='Name']"));

                string expectedErrorMessage = "The field Tên sản phẩm must be a string with a maximum length of 50.";
                string actualErrorMessage = validationErrorElement.Text;

                if (actualErrorMessage.Equals(expectedErrorMessage))
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Name";
                package.Save();
            #endregion
        }
        
        #endregion

        #region test category
            for (int i = 0; i < cateSelections.Length; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
            
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[i]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[i];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Category";
                package.Save();
            #endregion
        }
        

        #endregion
        
        #region test description
        
        //empty des
        for (int i = 0; i < numTest; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            var desElement = driver.FindElement(By.Id("Description"));
                desElement.Click();
                desElement.SendKeys("");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = ""; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                try
                {
                    var isValid =
                        (bool)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].checkValidity();",
                            desElement);
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                catch (WebDriverException e)
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                    
                }
                worksheet.Cells[row, 13].Value = "Description";
                package.Save();
            #endregion
        }
        //legal random des
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            int length = random.Next(1, 50);
            string desStr = GenerateRandomSpecialString(length);
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys(desStr);
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = desStr; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Description";
                package.Save();
            #endregion
        }
        
        #endregion

        #region test provider
            for (int i = 0; i < providerSelections.Length; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
            
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[i]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("100000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[i];
                worksheet.Cells[row, 6].Value = "100000"; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Provider";
                package.Save();
            #endregion
        }
        

        #endregion
        
        #region test price
        
        //empty price
        for (int i = 0; i < numTest; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            var priceelement = driver.FindElement(By.Id("Price"));
            priceelement.Click();
            priceelement.SendKeys("");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = ""; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                try
                {
                    var isValid =
                        (bool)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].checkValidity();",
                            priceelement);
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                catch (WebDriverException e)
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                    
                }
                worksheet.Cells[row, 13].Value = "Price";
                package.Save();
            #endregion
        }
        //legal random normal price
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            int length = random.Next(1, 10);
            string priceStr = GenerateRandomNumberString(length);
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys(priceStr);
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = priceStr; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Price";
                package.Save();
            #endregion
        }        
        // random negative price
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            int length = random.Next(1, 10);
            string priceStr = "-" + GenerateRandomNumberString(length);
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys(priceStr);
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("10");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = priceStr; 
                worksheet.Cells[row, 7].Value = "10";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Fail (price cant < 0)";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                worksheet.Cells[row, 13].Value = "Price";
                package.Save();
            #endregion
        }
        
        #endregion
        
        #region test quantity
        
        //empty quan
        for (int i = 0; i < numTest; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            var quantityElement = driver.FindElement(By.Id("Quantity"));
            quantityElement.Click();
            quantityElement.SendKeys("");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                try
                {
                    var isValid =
                        (bool)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].checkValidity();",
                            quantityElement);
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                catch (WebDriverException e)
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                    
                }
                worksheet.Cells[row, 13].Value = "Quantity";
                package.Save();
            #endregion
        }
        //legal random normal quantity
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price

            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            int length = random.Next(1, 10);
            string quanstr = GenerateRandomNumberString(length);
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys(quanstr);
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = quanstr;
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Quantity";
                package.Save();
            #endregion
        }        
        // random negative quantity
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price

            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            int length = random.Next(1, 10);
            string quanstr = "-" + GenerateRandomNumberString(length);
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys(quanstr);
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = quanstr;
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Quantity";
                package.Save();
            #endregion
        }
        #endregion
        
        #region test discount
        
        //empty discount
        for (int i = 0; i < numTest; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("50");
            #endregion

            #region Discount
            var discountElement = driver.FindElement(By.Id("Discount"));
            discountElement.Click();
            discountElement.SendKeys("");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "50";
                worksheet.Cells[row, 8].Value = ""; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                try
                {
                    var isValid =
                        (bool)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].checkValidity();",
                            discountElement);
                    worksheet.Cells[row, 12].Value = "Fail (discount can null)";
                }
                catch (WebDriverException e)
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                    
                }
                worksheet.Cells[row, 13].Value = "Discount";
                package.Save();
            #endregion
        }
        //legal random normal discount
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price

            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("50");
            #endregion

            #region Discount
            int length = random.Next(1, 100);
            string discountstr = length.ToString();
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys(discountstr);
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "50";
                worksheet.Cells[row, 8].Value = discountstr; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Discount";
                package.Save();
            #endregion
        }        
        // random negative discount
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price

            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity

            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("50");
            #endregion

            #region Discount
            int length = random.Next(-99, 0);
            string discountstr = length.ToString();
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys(discountstr);
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "50";
                worksheet.Cells[row, 8].Value = discountstr; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                worksheet.Cells[row, 13].Value = "Discount";
                package.Save();
            #endregion
        }
        
        // >100 discount
        for (int i = 0; i < numTest; i++)
        {
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name


                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price

            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity

            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("50");
            #endregion

            #region Discount
            int length = random.Next(100, 200);
            string discountstr = length.ToString();
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys(discountstr);
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            driver.FindElement(By.Id("UploadImage"))
                .SendKeys(imgpath[0]);
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;
                
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "50";
                worksheet.Cells[row, 8].Value = discountstr; 
                worksheet.Cells[row, 9].Value = imgpath[0]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                {
                    worksheet.Cells[row, 12].Value = "Fail";
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Pass";
                }
                worksheet.Cells[row, 13].Value = "Discount";
                package.Save();
            #endregion
        }
        #endregion
        
        #region test img
        
        //empty discount
        for (int i = 0; i < imgpath.Length; i++)
        {
            //empty
            #region  start test case
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:58473/Admin/Product/Add");
            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            #endregion

            #region name
                var nameElement = driver.FindElement(By.Id("Name"));
                nameElement.Click();
                driver.FindElement(By.Id("Name")).SendKeys("name");
            #endregion

            

            #region category
            driver.FindElement(By.Id("CateId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("CateId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{cateSelections[0]}']"))).Click();
            }
            #endregion

            #region description
            driver.FindElement(By.Id("Description")).Click();
            driver.FindElement(By.Id("Description")).SendKeys("dess");
            #endregion

            #region provider
            driver.FindElement(By.Id("ProviderId")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ProviderId"));
                dropdown.FindElement(By.XPath(string.Format($"//option[. = '{providerSelections[0]}']"))).Click();
            }
            #endregion

            #region price
            driver.FindElement(By.Id("Price")).Click();
            driver.FindElement(By.Id("Price")).SendKeys("10000");
            #endregion

            #region quantity
            driver.FindElement(By.Id("Quantity")).Click();
            driver.FindElement(By.Id("Quantity")).SendKeys("50");
            #endregion

            #region Discount
            driver.FindElement(By.Id("Discount")).Click();
            driver.FindElement(By.Id("Discount")).SendKeys("30");
            #endregion

            #region Upload Img
            //driver.FindElement(By.Id("UploadImage")).Click();
            if (i < 3)
            {
                driver.FindElement(By.Id("UploadImage"))
                    .SendKeys(imgpath[i]);
            }
            #endregion

            #region start date (discount)
            driver.FindElement(By.Id("StartDate")).Click();
            driver.FindElement(By.Id("StartDate")).SendKeys("10/10/2017");
            #endregion

            #region end date
            driver.FindElement(By.Id("EndDate")).Click();
            driver.FindElement(By.Id("EndDate")).SendKeys("10/10/2025");
            #endregion

            //submit
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(2)")).Click();
            
            #region check result

                var worksheet = package.Workbook.Worksheets[0]; 
                int row = worksheet.Dimension.End.Row + 1;

                // Ghi các giá trị nhập vào
                worksheet.Cells[row, 1].Value = "Test " + row;
                worksheet.Cells[row, 2].Value = "name";
                worksheet.Cells[row, 3].Value = cateSelections[0];
                worksheet.Cells[row, 4].Value = "dess"; 
                worksheet.Cells[row, 5].Value = providerSelections[0];
                worksheet.Cells[row, 6].Value = "10000"; 
                worksheet.Cells[row, 7].Value = "50";
                worksheet.Cells[row, 8].Value = "30"; 
                worksheet.Cells[row, 9].Value = imgpath[i]; 
                worksheet.Cells[row, 10].Value = "10/10/2017";
                worksheet.Cells[row, 11].Value = "10/10/2025"; 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                if (i <= 1)
                {
                    if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                    {
                        worksheet.Cells[row, 12].Value = "Pass";
                    }
                    else
                    {
                        worksheet.Cells[row, 12].Value = "Fail";
                    }
                }
                else
                {
                    if (driver.Url == "http://localhost:58473/Admin/Product/Show")
                    {
                        worksheet.Cells[row, 12].Value = "Fail";
                    }
                    else
                    {
                        worksheet.Cells[row, 12].Value = "True";
                    }
                }
                worksheet.Cells[row, 13].Value = "Discount";
                package.Save();
            #endregion
        }
        
        #endregion
        driver.Close();
        package.Dispose();
        MessageBox.Show("Chạy test xong, xem kết quả tại :" + filePath, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    static ExcelPackage? package;    
    static string filePath;
    public static void GenerateResultFile()
    {
        string fileName = $"TestResult{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

        filePath = Path.Combine( "F:\\Learn\\TMDT\\websitebandodunghoctap\\WebsiteDoDungHocTap\\TestingProgram\\Result\\" , fileName);   

        try
        {
            package = new ExcelPackage(new FileInfo(filePath));

            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "cate";
            worksheet.Cells[1, 4].Value = "description"; 
            worksheet.Cells[1, 5].Value = "provider";
            worksheet.Cells[1, 6].Value = "Price"; 
            worksheet.Cells[1, 7].Value = "quantity";
            worksheet.Cells[1, 8].Value = "discount"; 
            worksheet.Cells[1, 9].Value = "img path"; 
            worksheet.Cells[1, 10].Value = "Start date";
            worksheet.Cells[1, 11].Value = "end date"; 
            worksheet.Cells[1, 12].Value = "Result"; 

            package.Save();

            Console.WriteLine($"Excel file created successfully: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating Excel file: {ex.Message}");
        }
    }

    public static string GenerateRandomNormalString(int length)
    {
        StringBuilder result = new StringBuilder(length);
        Random random = new Random();

        for (int i = 0; i < length; i++)
        {
            result.Append(charSet[random.Next(charSet.Length)]);
        }
        return result.ToString();
    }    
    public static string GenerateRandomSpecialString(int length)
    {
        StringBuilder result = new StringBuilder(length);
        Random random = new Random();

        for (int i = 0; i < length; i++)
        {
            result.Append(charSetWithSpecialChar[random.Next(charSetWithSpecialChar.Length)]);
        }
        
        return result.ToString();
        
    }
    public static string GenerateRandomNumberString(int length)
    {
        StringBuilder result = new StringBuilder(length);
        Random random = new Random();

        for (int i = 0; i < length; i++)
        {
            result.Append(random.Next(10).ToString());
        }
        return result.ToString();
    }
}
using System;
using OpenQA.Selenium;
using System.Threading;
using NUnit.Framework;
using Npgsql;
using Dapper;
using OpenQA.Selenium.Support.UI;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Collections.Generic;

namespace NorthGate.Test.UI.UITest
{
    [TestFixture]
    public class BasicDataProcess : BaseSet
    {
        [OneTimeSetUp]
        public void TestMethod1()
        {
            try
            {
                driver.FindElement(By.Id("navbarUserDropdownMenuLink"));
                driver.Manage().Window.Maximize();
            }
            catch
            {
                base.Login();
                driver.Manage().Window.Maximize();
            }
        }
        List<TestResult> TestResult = new List<TestResult>();
        [Test]
        public void CheckVendor()
        {
            //驗證廠商資料
            driver.Url = "http://192.168.32.143:8080/admin/vendors";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            driver.FindElement(By.Id("vendorCode")).SendKeys("MIC01");
            driver.FindElement(By.Id("description")).SendKeys("妙承");
            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("div#swal2-html-container")).Text.Contains("此「廠商代碼」已經存在。");
            var data = new TestResult { CheckItem = "廠商資料", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
            
        }
        [Test]
        public void CheckDrivers()
        {
            //驗證司機資料
            driver.Url = "http://192.168.32.143:8080/admin/drivers";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            driver.FindElement(By.Id("driverName")).SendKeys("UI測試司機");
            driver.FindElement(By.Id("appAccount")).SendKeys("Testdriver");
            driver.FindElement(By.Id("appPassword")).SendKeys("Testdriver");
            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("h2#swal2-title")).Text.Contains("新增理貨資料成功");
            var data = new TestResult { CheckItem = "司機資料", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }

        [Test]
        public void CheckTallyfreights()
        {
            //驗證理貨明細
            driver.Url = "http://192.168.32.143:8080/admin/tally-freights";
            Thread.Sleep(1500);
            driver.FindElement(By.Id("ladingBillNo")).SendKeys("0000");
            driver.FindElement(By.Id("serialNumber")).SendKeys("0000");
            driver.FindElement(By.Id("customerPhoneNo")).SendKeys("0912345678");
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("scrollBy(0,400);");
            Thread.Sleep(1500);
            driver.FindElement(By.XPath(".//button[text()='搜尋']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("tbody tr td.nrg-data-table-column:nth-of-type(3)")).Text.Contains("D1");
            var data = new TestResult { CheckItem = "理貨明細", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }

        [Test]
        public void CheckTallysheets()
        {
            //驗證理貨單
            driver.Url = "http://192.168.32.143:8080/admin/tally-sheets";
            Thread.Sleep(1000);
            driver.FindElement(By.Id("ladingBillNo")).SendKeys("0000");
            driver.FindElement(By.Id("customerName")).SendKeys("Test");
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("scrollBy(0,500);");
            Thread.Sleep(1000);
            driver.FindElement(By.XPath(".//button[text()='搜尋']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("tbody tr td.nrg-data-table-column:nth-of-type(2)")).Text.Contains("2022/06/24");
            var data = new TestResult { CheckItem = "理貨單", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }
        [Test]
        public void CheckExceptionReasons()
        {
            //驗證例外原因管理
            driver.Url = "http://192.168.32.143:8080/admin/exception-reasons/";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            driver.FindElement(By.Id("reasonName")).SendKeys("測試狀況");
            driver.FindElement(By.Id("description")).SendKeys("測試狀況");
            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("h2#swal2-title")).Text.Contains("新增例外原因資料成功");
            var data = new TestResult { CheckItem = "例外原因管理", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }
        [Test]
        public void CheckDeliveryAreas()
        {
            //驗證派送區域代碼
            driver.Url = "http://192.168.32.143:8080/admin/delivery-areas";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            driver.FindElement(By.Id("code")).SendKeys("UItest");
            driver.FindElement(By.Id("description")).SendKeys("UI測試");
            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("h2#swal2-title")).Text.Contains("新增派送區域資料成功");
            var data = new TestResult { CheckItem = "派送區域代碼", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }

        [Test]
        public void CheckDeliveryAreaBindings()
        {
            //驗證派送區域綁定
            driver.Url = "http://192.168.32.143:8080/admin/delivery-area-bindings";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            Thread.Sleep(1000);
            SelectElement oSelect = new SelectElement(driver.FindElement(By.Id("delivery_area_id")));
            oSelect.SelectByText("D1 (松山區)");
            driver.FindElement(By.CssSelector("div.css-1s2u09g-control")).Click();
            driver.FindElement(By.XPath("//*[text()='臺北市']")).Click();
            driver.FindElement(By.CssSelector("div.css-1s2u09g-control")).Click();
            driver.FindElement(By.XPath("//*[text()='松山區']")).Click();

            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("h2#swal2-title")).Text.Contains("新增區域綁定資料成功");
            var data = new TestResult { CheckItem = "派送區域綁定", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }

        [Test]
        public void CheckOtherFeeTitles()
        {
            //驗證其他費用名目設定
            driver.Url = "http://192.168.32.143:8080/admin/other-fee-titles";
            driver.FindElement(By.XPath(".//button[text()='新增']")).Click();
            driver.FindElement(By.Id("title")).SendKeys("UI測試");
            driver.FindElement(By.XPath("//*[@type='submit']")).Click();
            Thread.Sleep(1000);
            bool result = driver.FindElement(By.CssSelector("h2#swal2-title")).Text.Contains("新增其他費用名目資料成功");
            var data = new TestResult { CheckItem = "其他費用名目設定", Result = result };
            TestResult.Add(data);
            Assert.IsTrue(result);
        }

        [OneTimeTearDown]
        public void cleanup()
        {
            //string filepath = @"C:\Users\essences\Desktop\廠商帳款報表.xlsx";
            //XSSFWorkbook workbook = new XSSFWorkbook(File.Open(filepath, FileMode.Open));
            //var sheet = workbook.GetSheetAt(0);
            //string[] test = new string[sheet.LastRowNum];
            //for (var num = 0; num < sheet.LastRowNum; num++)
            //{
            //    test[num] = sheet.GetRow(num).GetCell(4).StringCellValue.Trim();
            //}
            
            //建立測試報告Excel
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet1");

            sheet.CreateRow(0);
            sheet.GetRow(0).CreateCell(0).SetCellValue("檢查項目");
            sheet.GetRow(0).CreateCell(1).SetCellValue("檢查結果");
            for (var num = 0; num < TestResult.Count; num++)
            {
                sheet.CreateRow(num + 1);
                sheet.GetRow(num+1).CreateCell(0).SetCellValue(TestResult[num].CheckItem);
                sheet.GetRow(num+1).CreateCell(1).SetCellValue(TestResult[num].Result?"PASS":"FAIL");
            }

            FileStream file = new FileStream(@"C:\Users\essences\Desktop\UITest\北門UI測試-" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx", FileMode.Create, System.IO.FileAccess.ReadWrite);
            workbook.Write(file);
            file.Close();
            workbook.Close();
            //刪除測試資料
            var connstr = "Server=192.168.32.143;Database=north_gate;Username=north_gate_admin;Password=@$&&%@%&essens;";
            NpgsqlConnection connection = new NpgsqlConnection(connstr);

            connection.Open();
            string DeleteTestDriver = @"DELETE from driver where driver_name = 'UI測試司機'";
            string DeleteTestException = @"DELETE from exception_reason where reason_name = '測試狀況'";
            string DeleteDelivertArea = @"DELETE from delivery_area where code = 'UItest'";
            string DeleteOtherFeeTitles = @"DELETE from other_fee_title where title = 'UI測試'";
            Guid DeleteDelivertAreaBingingsId = connection.QueryFirstOrDefault<Guid>(@"SELECT delivery_area_binding_id from delivery_area_binding where district = '松山區' and county = '臺北市' order by created_at DESC limit 1");
            string DeleteDelivertAreaBingings = "DELETE from delivery_area_binding where delivery_area_binding_id = @DeleteDelivertAreaBingingsId";
            connection.Execute(DeleteTestDriver);
            connection.Execute(DeleteTestException);
            connection.Execute(DeleteDelivertArea);
            connection.Execute(DeleteOtherFeeTitles);
            connection.Execute(DeleteDelivertAreaBingings, new { DeleteDelivertAreaBingingsId });
            connection.Close();

            //關閉瀏覽器
            driver.Quit();
        }

     }
    public class TestResult
    {
        public bool Result { get; set; }
        public string CheckItem { get; set; }
    }
}

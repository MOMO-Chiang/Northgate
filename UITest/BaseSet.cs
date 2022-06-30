using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace NorthGate.Test.UI.UITest
{
    public class BaseSet
    {
        
        protected IWebDriver driver;
        public BaseSet()
        {
            //指定Chrome驅動程式(須隨google版本更新)
            this.driver = new ChromeDriver(@"C:\Users\essences\Downloads\chromedriver_win32");
        }
        protected void Login()
        {
            driver.Navigate().GoToUrl("http://192.168.32.143:8080/admin/login");
            IWebElement Account = driver.FindElement(By.XPath("//*[@type='text']"));
            Account.SendKeys("Test");
            IWebElement Password = driver.FindElement(By.XPath("//*[@type='password']"));
            Password.SendKeys("1234");
            IWebElement Submit = driver.FindElement(By.XPath("//*[@type='submit']"));
            Submit.Click();
        }

    }
}

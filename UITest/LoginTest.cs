using System;
using NUnit.Framework;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace NorthGate.Test.UI.UITest
{
    [TestFixture]
    public class LoginTest : BaseSet
    {
        [Test]
        public void TestMethod1()
        {
            base.Login();
            driver.Quit();
        }
    }
}

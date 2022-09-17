using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;

namespace AutomationAssignment
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            IWebDriver driver = new EdgeDriver();
            driver.Manage().Window.Maximize();
            Functions function = new Functions(driver);
            function.OpenBrowser();
            function.ActionFunction();
            function.CloseBrowser();
        }
    }
}

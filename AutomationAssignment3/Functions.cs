using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Edge;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;


namespace AutomationAssignment
{

    public class Functions
    {
        IWebDriver driver;
        static string path = Path.Combine("DataEX.xlsx");      //Located in AutomationAssignment\bin\Debug\netcoreapp3.1
        XSSFWorkbook workbook = new XSSFWorkbook(File.Open(path, FileMode.Open));
        string url = System.IO.File.ReadAllText("Url.txt");    //Located in AutomationAssignment\bin\Debug\netcoreapp3.1

        public Functions(IWebDriver driver)
        {
            this.driver = driver;
        }

        public void OpenBrowser()
        {
            driver.Navigate().GoToUrl(url);
        }

        public void ActionFunction()
        {
            var sheet = workbook.GetSheetAt(0);
            int rows = sheet.LastRowNum;

            for (int i = 1; i < rows+1; i++)
            {
                DataFormatter formatter = new DataFormatter();
                int waittime =int.Parse(formatter.FormatCellValue(sheet.GetRow(i).GetCell(1)).Trim());
                string cssselectorName = formatter.FormatCellValue(sheet.GetRow(i).GetCell(2)).Trim();
                var op = formatter.FormatCellValue(sheet.GetRow(i).GetCell(3)).Trim();
                var value = formatter.FormatCellValue(sheet.GetRow(i).GetCell(4)).Trim();

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waittime));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector(cssselectorName)));
                IWebElement element = driver.FindElement(By.CssSelector(cssselectorName));

               

                try
                {

                    switch (op)
                    {
                        case "Click":
                            element.Click();
                            break;

                        case "Set":
                           
                            element.SendKeys(value);
                            break;

                        case "Displayed":

                            bool displayed = element.Displayed;
                            if (displayed == false)
                            {
                                throw new Exception("Element " + cssselectorName + " not found");
                            }
                            break;

                    }
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message + "Canoot find selector "+cssselectorName);
                }
            }
        }
        public void CloseBrowser()
        {
            driver.Close();
            driver.Quit();
        }
    }

}

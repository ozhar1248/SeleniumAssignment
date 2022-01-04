using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace AssignmentOz
{
    class Program
    {
        static Microsoft.Office.Interop.Excel.Application xlApp;
        static Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
        static Microsoft.Office.Interop.Excel._Worksheet xlWorksheet;
        static Microsoft.Office.Interop.Excel.Range xlRange;
        static Dictionary<string, string> hashVarData; // hash table that maps variable name in column 1 from excel file
                                                // to the value that appears in column 2

        static void ReadDataToArray()
        {
            
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            string path = Directory.GetCurrentDirectory();
            xlWorkbook = xlApp.Workbooks.Open(path+"\\data.xlsx");
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    hashVarData.Add(xlRange.Cells[i, 1].Value2.ToString(), xlRange.Cells[i, 2].Value2.ToString());
                    //myAL.Add(xlRange.Cells[i, 2].Value2.ToString());
                }
            }

        }
        static void CloseData()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        static WebDriver Login()
        {
            WebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(hashVarData["URL"]);
            driver.FindElement(By.Id(hashVarData["ID_USERNAME"])).SendKeys(hashVarData["USERNAME"]);
            driver.FindElement(By.Id(hashVarData["ID_PASSWORD"])).SendKeys(hashVarData["PASSWORD"]);
            driver.FindElement(By.Id(hashVarData["ID_BUTTON_LOGIN"])).Click();
            return driver;
        }
        static void CheckLoginTime(WebDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Int32.Parse(hashVarData["TIME_TO_LOAD_SECONDS"])));
            long time1 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(hashVarData["ID_ACCOUNT_LINK"])));
            long time2 = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            if (time2-time1 >= Int32.Parse(hashVarData["TIME_TO_LOAD_SECONDS"]) * 1000)
            {
                throw new Exception("time of login is over 10 seconds");
            }
        }

        static void CheckLogout(WebDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Int32.Parse(hashVarData["TIME_TO_LOAD_SECONDS"])));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(hashVarData["ID_ACCOUNT_LINK"])));

            driver.FindElement(By.Id(hashVarData["ID_ACCOUNT_LINK"])).Click();
            driver.FindElement(By.Id(hashVarData["ID_LOGOUT"])).Click();
            driver.FindElement(By.Id(hashVarData["ID_OK_LOGOUT"])).Click();

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Int32.Parse(hashVarData["TIME_TO_LOAD_SECONDS"])));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(hashVarData["SELECTOR_HOVER"])));
            string text = driver.FindElement(By.CssSelector(hashVarData["SELECTOR_HOVER"])).GetAttribute("title");
            string path = Directory.GetCurrentDirectory();
            using (StreamWriter writer = new StreamWriter(path+hashVarData["FILE_NAME"]))
            {
                writer.WriteLine(text);
            }
        }

        static void Main(string[] args)
        {
            hashVarData = new Dictionary<string, string>();
            ReadDataToArray();
            CloseData();
            WebDriver driver = Login();
            CheckLoginTime(driver);
            CheckLogout(driver);

        }
    }
}

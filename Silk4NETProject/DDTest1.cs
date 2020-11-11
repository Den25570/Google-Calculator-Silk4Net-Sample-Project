using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SilkTest.Ntf;
using SilkTest.Ntf.XBrowser;

namespace Silk4NETProject
{
    [SilkTestClass]
    public class DDTest1
    {
        private readonly Desktop _desktop = Agent.Desktop;

        [TestInitialize]
        public void Initialize()
        {
            // Go to web page 'https://www.google.com/search?q=%D0%BA%D0%B0%D0%BB%D1%8C%D0%BA%D1%83%D0%BB%D1%8F%D1%82%D0%BE%D1%80&rlz=1C1CHBD_enBY927BY927&oq=%D0%BA%D0%B0%D0%BB%D1%8C%D0%BA%D1%83%D0%BB%D1%8F%D1%82%D0%BE%D1%80&aqs=chrome..69i57j0l7.4823j0j7&sourceid=chrome&ie=UTF-8'
            BrowserBaseState baseState = new BrowserBaseState();
            baseState.Execute();
        }

        [TestMethod]
        public void TestMethod1()
        {
            BrowserApplication webBrowser = _desktop.BrowserApplication("google_com");
            BrowserWindow browserWindow = webBrowser.BrowserWindow("BrowserWindow");     

            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(@"d:\Projects\SilkTestLab5\Silk4NETProject\DDT.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); ;
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1); ;
            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.UsedRange;
            int rw = 2;
            for (int rCnt = 1; rCnt <= rw; rCnt++)
            {
                string val1 = (range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                string operation = (range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                string val2 = (range.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                string res = (range.Cells[rCnt, 4] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();

                browserWindow.DomElement(val1).Click();
                browserWindow.DomElement(operation).Click();
                browserWindow.DomElement(val2).Click();
                browserWindow.DomElement("DIV").Click();
                Assert.AreEqual(res, browserWindow.DomElement("cwos").GetProperty("textContents"));
                Assert.AreEqual(res, browserWindow.DomElement("cwos").Text);
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            browserWindow.Close();
        }
    }
}
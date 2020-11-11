using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SilkTest.Ntf;
using SilkTest.Ntf.XBrowser;

namespace Silk4NETProject
{
    [SilkTestClass]
    public class Test1
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
            browserWindow.DomElement("4").Click();
            browserWindow.DomElement("+").Click();
            browserWindow.DomElement("8").Click();
            browserWindow.DomElement("DIV").Click();
            browserWindow.DomElement("cwos").Click();
            browserWindow.DomElement("AC").Click();
            Assert.AreEqual("0", browserWindow.DomElement("cwos").Text);
            browserWindow.Close();
        }
    }
}
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using NUnit.Framework;
using System;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;


namespace SeleniumDotNetHelpLibrary.Selenium.Support.Driver
{
    public class Driver
    {
        public IWebDriver driver;

        public IWebDriver driverIE()
        {
            var options = new InternetExplorerOptions();
            //options.ForceCreateProcessApi = true;
            //options.BrowserCommandLineArguments = "-private";

            //options.AddAdditionalCapability("ignoreZoomSetting", true);
            options.EnableNativeEvents = false;
            options.UnexpectedAlertBehavior = InternetExplorerUnexpectedAlertBehavior.Ignore;
            //options.AddAdditionalCapability("ignoreProtectedModeSettings", true);
            //options.AddAdditionalCapability("disable-popup-blocking", true);
            options.EnablePersistentHover = true;
            //options."ACCEPT_SSL_CERTS", true);
            options.RequireWindowFocus = true;


            options.EnsureCleanSession = true;
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            options.IgnoreZoomLevel = true;
            driver = new InternetExplorerDriver(options);
            driver.Manage().Window.Maximize();
            return driver;
        }

        public IWebDriver driverChrome()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            return driver;
        }
    
        public IWebDriver driverFirefox()
        {
            driver = new FirefoxDriver();
            driver.Manage().Window.Maximize();
            return driver;
        }

    }
}

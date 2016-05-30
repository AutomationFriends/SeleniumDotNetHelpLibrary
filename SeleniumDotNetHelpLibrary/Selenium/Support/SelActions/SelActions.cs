using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using NUnit.Framework;
using System;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using OpenQA.Selenium.Interactions;

namespace SeleniumDotNetHelpLibrary.Selenium.Support.SelActions
{

    public class SelActions
    {
        public IWebDriver driver;

        public SelActions(IWebDriver driver)
        {
            this.driver = driver;
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5));
            PageFactory.InitElements(driver, this);
        }

        public SelActions()
        { }

        // To wait:

        public static void WaitForElement(IWebElement element, int timeToWaitInSeconds = 10, int pollingIntervalInMilliSeconds = 1000)
        {
            for (int i = 0; i < timeToWaitInSeconds; i++)
            {
                if (element.Displayed)
                {
                    break;
                }
                Thread.Sleep(pollingIntervalInMilliSeconds);
            }
        }

        public static void WaitForElements(IList<IWebElement> elements, int timeToWaitInSeconds = 1, int pollingIntervalInMilliSeconds = 1000)
        {
            for (int i = 0; i < timeToWaitInSeconds; i++)
            {
                if (elements.Count.Equals(1))
                {
                    break;
                }
                Thread.Sleep(pollingIntervalInMilliSeconds);
            }
        }

        public static void WaitForPageLoad(int maxWaitTimeInSeconds, IWebDriver driver)
        {

            string state = string.Empty;
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

                //Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns true
                wait.Until(d =>
                {

                    try
                    {
                        state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                    }
                    catch (InvalidOperationException)
                    {
                        //Ignore
                    }
                    catch (NoSuchWindowException)
                    {
                        //when popup is closed, switch to last windows
                        //driver.SwitchTo().Window();
                    }
                    //In IE7 there are chances we may get state as loaded instead of complete
                    return (state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase));

                });
            }
            catch (TimeoutException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (NullReferenceException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (WebDriverException)
            {
                if (driver.WindowHandles.Count == 1)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }
                state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
                    throw;
            }
        }

        public static void WaitAndSwitchToAlert(int maxWaitTimeInSeconds, IWebDriver driver)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

            wait.Until(d =>
            {

                try
                {
                    driver.SwitchTo().Alert();
                    return true;
                }

                catch (NoAlertPresentException ex)
                {
                    throw new NoAlertPresentException("Something wrong:", ex);
                }

            });

        }

        public static void WaitSwitchToAlertAndAccept(int maxWaitTimeInSeconds, IWebDriver driver)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

            wait.Until(d =>
            {

                try
                {
                    driver.SwitchTo().Alert().Accept();
                    return true;
                }

                catch (NoAlertPresentException ex)
                {
                    throw new NoAlertPresentException("Something wrong:", ex);
                }

            });


        }

        public static void WaitForAjax(IWebDriver driver, int timeoutSecs, bool throwException)
        {
            for (var i = 0; i < timeoutSecs; i++)
            {
                var ajaxIsComplete = (bool)(driver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete) return;
            }
            if (throwException)
            {
                throw new Exception("WebDriver timed out waiting for AJAX call to complete");
            }
        }

        // To write: 

        public static void writeTextToField(string inputText, IWebElement element)
        {
            element.Clear();
            element.SendKeys(inputText);
            Thread.Sleep(200);
        }

        // Manage windows:

        public static string RememberMyCurrentWindow(IWebDriver driver)
        {
            string currentWindow = driver.CurrentWindowHandle;
            return currentWindow;
        }

        public static void SwitchToWindow(string currentWindow, string titleNewWindow, IWebDriver driver)
        {

            WaitForPageLoad(5, driver);

            ReadOnlyCollection<string> handles = driver.WindowHandles;

            foreach (string handle in handles)
            {

                if (handle != currentWindow)
                {
                    if (driver.SwitchTo().Window(handle).Title.Contains(titleNewWindow))
                        break;
                }
            }
        }

        // Working with pages:

        public static void verifyPageTitle(string pageTitle, IWebDriver driver)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(5000)).Until(ExpectedConditions.ElementExists((By.TagName("title"))));
            if (driver.Title != pageTitle)
                throw new NoSuchWindowException("This is not the: " + pageTitle + " page" + " This is: " + driver.Title);
        }

        // Working with drop-downmooo

        public static void selectFromDropDownByText(string inputText, IWebElement dropDown)
        {
            SelectElement list = new SelectElement(dropDown);


            IList<IWebElement> allElements = list.Options;

            foreach (IWebElement element in allElements)
            {
                if (element.Text.Contains(inputText))
                {
                    element.Click();
                }

            }
        }

        public string getTextSelectedInDropDown(IWebElement dropDown)
        {
            SelectElement select = new SelectElement(dropDown);
            var option = select.SelectedOption;
            string selectedElement = option.Text;

            return selectedElement;
        }

        // Working with the similar objects

        public IWebElement searchElementByTextInList(string textInAlert, IList<IWebElement> allElements)
        {
            IWebElement myAlert = null;

            foreach (IWebElement element in allElements)
            {
                string text = element.Text.Replace("\"", "");

                if (text.Equals(textInAlert))
                {
                    myAlert = element;
                }
            }

            return myAlert;
        }

        public IWebElement searchElementByText(string elementHasText, IList<IWebElement> allElements)
        {
            IWebElement myElement = null;


            foreach (IWebElement element in allElements)
            {

                if (element.Text.Contains(elementHasText))
                {
                    myElement = element;
                }
            }

            return myElement;

        }

        public void clickOnElementIfItIsClickable(IList<IWebElement> allElements)
        {
            IWebElement myElement = null;

            try
            {

                foreach (IWebElement element in allElements)
                {

                    if (element.Enabled)
                    {
                        if (element.Displayed)
                        {
                            element.Click();
                            break;               
                        }

                    }

                }
            }

            catch (NoSuchElementException)
            {

            }

        }
    }
}

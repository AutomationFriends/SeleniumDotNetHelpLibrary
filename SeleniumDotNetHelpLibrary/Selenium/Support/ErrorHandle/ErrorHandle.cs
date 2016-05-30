using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumDotNetHelpLibrary.Selenium.Support.ErrorHandle
{
    public class ErrorHandle
    {
        public void takeScreenshot(IWebDriver driver)
        {
            if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
            {
                ITakesScreenshot ssdriver = driver as ITakesScreenshot;
                Screenshot screenshot = ssdriver.GetScreenshot();
                string projectPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd-hhmm-ss");
                string testName = NUnit.Framework.TestContext.CurrentContext.Test.Name;
                testName = (testName.Substring(0, testName.LastIndexOf("(") + 1)).Replace("(", "");
                screenshot.SaveAsFile(projectPath + "\\Screenshots\\" + testName + "-" + timestamp + ".png", ImageFormat.Png);
            }
        }
    }
}

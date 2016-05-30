using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;


namespace SeleniumDotNetHelpLibrary.Selenium.Support.Logging
{
    public class Logging
    {
        public void SimpleLogging()
        {

            var st = new StackTrace();
            var sf = st.GetFrame(1);

            var className = sf.GetMethod().DeclaringType.Name;
            var methodName = sf.GetMethod().Name;
            var location = string.Format("{0}.{1}", className, methodName);

            //string method = string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name);
            string path = @"C:\temp\file.log";
            string appendText = "LOG:" + location + Environment.NewLine;
            File.AppendAllText(path, appendText);
        }

    }
}

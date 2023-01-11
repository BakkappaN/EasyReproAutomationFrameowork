using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System;
using System.Drawing;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Util
{
    [TestClass]
    public class TestBase
    {
        /* static public SecureString _username;
         static public SecureString _password;
         static public SecureString _mfaSecretKey;

         public static Uri _xrmUri;

         public static string parentWindow;
         public static WebClient client;
         public static XrmApp xrmApp;*/

        public static XrmPage xrmPage;
        public static TestContext testContextInstance;
        public static WebClient client;
        public static XrmApp xrmApp;

        public TestContext TestContext
        {
            get { return testContextInstance; }
            set { testContextInstance = value; }
        }
        /*  public static IWebDriver _driver;

          public static TestContext testContextInstance;

          public TestContext TestContext
          {
              get { return testContextInstance; }
              set { testContextInstance = value; }
          }

          [AssemblyInitialize]
          public static void AssemblyInitialize(TestContext context)
          {
              Console.WriteLine("Assembly Initilization...");
              testContextInstance = context;
              _username = testContextInstance.Properties["OnlineUsername"].ToString().ToSecureString();
              _password = testContextInstance.Properties["OnlinePassword"].ToString().ToSecureString();
              _mfaSecretKey = testContextInstance.Properties["MfaSecretKey"].ToString().ToSecureString();
          }

          [AssemblyCleanup]
          public static void AssemblyCleanUp()
          {
              Console.WriteLine("Assembly Cleanup...");
              _username.Clear();
              _password.Clear();
          }*/

        [TestInitialize]
        public void TestInitialize()
        {
            Console.WriteLine("Test started...");
            /*   client = new WebClient(TestSettings.Options);
               xrmApp = new XrmApp(client);
               xrmPage = new XrmPage(client.Browser);
               new Uri(testContextInstance.Properties["OnlineCrmUrl"].ToString());*/

            client = new WebClient(TestSettings.Options);
            xrmApp = new XrmApp(client);
            xrmPage = new XrmPage(client.Browser);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            Console.WriteLine("Test Cleanup...");

            //Take screenshot when test failed
            try
            {
                String status = testContextInstance.CurrentTestOutcome.ToString();
                if (status.Equals("Failed") || status.Equals("Error"))
                {
                    xrmPage.Browser.Driver.Manage().Window.Size = new Size(1920, 1080);
                }
                var currentDateandTime = testContextInstance.TestName + "_" + DateTime.Now.ToString("MM_dd_yyyy_hh_mm_ss");
                string fileName = string.Format("Failed - " + currentDateandTime + ".jpeg", testContextInstance.TestResultsDirectory);
                xrmPage.Browser.TakeWindowScreenShot(fileName, ScreenshotImageFormat.Jpeg);
                testContextInstance.AddResultFile(fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            xrmPage.Browser.Driver.Quit();
            xrmPage.Browser.ThinkTime(3000);
        }
    }
}

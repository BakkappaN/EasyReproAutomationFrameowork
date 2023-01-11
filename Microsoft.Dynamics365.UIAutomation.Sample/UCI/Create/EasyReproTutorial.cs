// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Sample.Util;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System.Drawing;
using OpenQA.Selenium;


namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class EasyReproTutorial : TestBase
    {
        // public TestContext TestContext { get; set; }

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        [TestCategory("TestersTalk")]
        public void EasyReproTest()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
            xrmApp.CommandBar.ClickCommand("New");
            //textbox
            xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));
            xrmApp.Entity.SetValue("telephone1", "1234567897");

            //Lookup
            xrmApp.Entity.SetValue("primarycontactid", "alex baker");
            xrmApp.Lookup.OpenRecord(0);

            //TABS
            xrmApp.Entity.SelectTab("Details");

            //Dropdown list
            xrmApp.Entity.SetValue(new OptionSet { Name = "ownershipcode", Value = "1" });
            xrmApp.ThinkTime(4000);
            xrmApp.Entity.Save();
        }

        [TestMethod]
        [TestCategory("SmokeTesting"), TestCategory("TestersTalk")]
        [Priority(0)]
        public void EasyReproTest2()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Contacts");
            xrmApp.CommandBar.ClickCommand("New");

            //TABS
            xrmApp.Entity.SelectTab("Details");

            var birthDate = new DateTimeControl("birthdate") { Value = DateTime.Now.AddDays(-1000) };
            xrmApp.Entity.SetValue(birthDate);
            xrmApp.ThinkTime(4000);
            xrmApp.Entity.Save();
        }

        [TestMethod]
        [TestCategory("TestersTalk")]
        [Priority(0)]
        public void EasyReproTest3()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
            xrmApp.CommandBar.ClickCommand("New");

            var accountName = TestSettings.GetRandomString(5, 15);
            xrmApp.Entity.SetValue("name", accountName);
            xrmApp.Entity.Save();

            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

            xrmApp.Grid.Search(accountName);

            xrmApp.ThinkTime(3000);
        }

        [TestMethod]
        [TestCategory("TestersTalk")]
        public void EasyReproTest4()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

            //Select view
            // xrmApp.Grid.SwitchView("Active Accounts");

            xrmApp.CommandBar.ClickCommand("New");

            var accountName = TestSettings.GetRandomString(5, 15);
            xrmApp.Entity.SetValue("name", accountName);
            xrmApp.Entity.Save();

            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

            xrmApp.Grid.Search(accountName);

            xrmApp.ThinkTime(3000);

            xrmApp.Grid.OpenRecord(0);

            xrmApp.CommandBar.ClickCommand("Delete");

            xrmApp.Dialogs.ConfirmationDialog(true);

            xrmApp.ThinkTime(7000);
        }

        [TestMethod]
        [TestCategory("TestersTalk")]
        public void EasyReproTest5()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

            xrmApp.CommandBar.ClickCommand("New");

            var accountName = TestSettings.GetRandomString(5, 15);
            xrmApp.Entity.SetValue("name", accountName);
            xrmApp.Entity.Save();

            xrmApp.Entity.SelectTab("Related", "Contacts");

            xrmApp.RelatedGrid.ClickCommand("New Contact");

            xrmApp.QuickCreate.SetValue("lastname", "TestersTalk");

            xrmApp.QuickCreate.Save();

            xrmApp.ThinkTime(4000);
        }

        [TestMethod]
        [TestCategory("TestersTalk")]
        [DeploymentItem("|DataDirectory|\\testdata.xml")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML",
        "|DataDirectory|\\testdata.xml",
         "testdata",
        DataAccessMethod.Sequential)]
        public void EasyReproDataDrivenTesting()
        {
            xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
            xrmApp.Navigation.OpenApp(UCIAppName.Sales);
            xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

            xrmApp.CommandBar.ClickCommand("New");

            xrmApp.Entity.SetValue("name", testContextInstance.DataRow["tag1"].ToString());
            xrmApp.Entity.Save();

            xrmApp.ThinkTime(4000);
        }

    }
}
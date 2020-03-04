using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BrightEdgeAutomationTool
{
    public static class PuppetMaster
    {
        public static IWebDriver Driver { get; set; }

        public static MainWindow Window { get; set; }

        private static string BaseURL = "https://app1.brightedge.com/ui/platform-r/instant/bulk_keyword_volume/";

        public static void LoginUser()
        {
            if (Driver.ElementExist(By.Name("data[User][login]")))
            {
                // Try to login
                var emailField = Driver.FindElement(By.Name("data[User][login]"), 20);
                emailField.SendKeys("john.connolly@galileotechmedia.com");

                var passwordField = Driver.FindElement(By.Name("data[User][password]"), 20);
                passwordField.SendKeys("Galileo123");

                var loginButton = Driver.FindElement(By.Id("login_submit"), 20);
                loginButton.Click();

                Thread.Sleep(1000);

                Driver.Navigate().GoToUrl(BaseURL);

            }
        }
    }
}

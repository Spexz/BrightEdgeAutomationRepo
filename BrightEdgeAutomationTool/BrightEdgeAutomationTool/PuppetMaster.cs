using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
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

        public static void RunProcess(string keywords, string location)
        {

            if (Driver.ElementExist(By.XPath("//div[a/button/span[contains(., 'All Queries')]]")))
            {
                var allQueries = Driver.FindElement(By.Id("//div[a/button/span[contains(., 'All Queries')]]"), 20);
                Click(allQueries);
            }

            var keywordField = Driver.FindElement(By.Id("keywords"), 20);
            keywordField.Clear();
            Click(keywordField);
            System.Windows.Clipboard.SetText(keywords);
            keywordField.SendKeys(OpenQA.Selenium.Keys.Control + "v");

            // Add location click
            var addLocation = Driver.FindElement(By.XPath("//div[contains(@data-testid, 'locations')]"), 20);
            Click(addLocation);

            // Input location
            var inputLocation = Driver.FindElement(By.XPath("//input[contains(@name, 'searchTerm')]"), 20);
            Click(inputLocation);
            System.Windows.Clipboard.SetText(location);
            inputLocation.SendKeys(OpenQA.Selenium.Keys.Control + "v");

            // Selection country checkbox
            var checkLocation = Driver.FindElement(By.XPath($"//div[input[@id= '{location}']]"), 20);
            Click(checkLocation);

            // collapse add location
            //xpath: //div[contains(@data-testid, 'locations')]
            var collapseLocation = Driver.FindElement(By.XPath("//div[contains(@data-testid, 'locations')]"), 20);
            Click(collapseLocation);

            var initialDetailsUrl = GetDetailsUrl();

            // Get volume button
            var getVolume = Driver.FindElement(By.XPath("//button[contains(@data-testid, 'queryButton')]"), 20);
            Click(getVolume);

            Console.WriteLine(initialDetailsUrl);

            RetryUntilSuccessOrTimeout(() => {
                try
                {
                    var detailsUrl = GetDetailsUrl();

                    if (String.IsNullOrEmpty(initialDetailsUrl) && !String.IsNullOrEmpty(detailsUrl))
                        return true;
                    else if (!initialDetailsUrl.Equals(detailsUrl) && detailsUrl != null)
                        return true;
                    else
                        return false;
                }
                catch (Exception e)
                {
                    return false;
                }

            }, TimeSpan.FromSeconds(60));

            //var rowElement = GetFirstRow();

            // Click the View Details button
            // //div[contains(@class, 'viewportRow')]//button
            var viewDetailsButton = Driver.FindElement(By.XPath("//div[contains(@class, 'viewportRow')]//button"), 20);
            Click(viewDetailsButton);

            // Optional: Click total search volume once to sort

            

            // Now download the csv
            var downloadButton = Driver.FindElement(By.XPath("//div[contains(@class, 'popover__target') and span/button[span[contains(text(), 'd')]]]"), 20);
            Click(downloadButton);
        }


        private static void Click(IWebElement element)
        {
            RetryUntilSuccessOrTimeout(() => {

                try
                {
                    Actions builder = new Actions(Driver);
                    //Actions hoverClick = builder.MoveToElement(element).MoveByOffset(1, 1).Click();
                    Actions hoverClick = builder.MoveToElement(element).Click();
                    hoverClick.Build().Perform();
                    
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Clicking: " + e.Message);
                    return false;
                }

            }, TimeSpan.FromSeconds(10));
        }

        public static bool RetryUntilSuccessOrTimeout(Func<bool> task, TimeSpan timeSpan)
        {
            bool success = false;
            int elapsed = 0;
            while ((!success) && (elapsed < timeSpan.TotalMilliseconds))
            {
                Thread.Sleep(1000);
                elapsed += 1000;
                success = task();
            }
            return success;
        }

        private static string GetDetailsUrl()
        {
            var scriptStr =
                "try {" +
                "return document.querySelectorAll('[class^=\"viewportRow\"]')[0]" +
                ".querySelector('[class^=\"cell\"] a[href]').getAttribute(\"href\");" + 
                "} catch(e) { return null; }";

            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)Driver;
            var url = (string)javaScriptExecutor.ExecuteScript(scriptStr);

            return url;
        }

        private static IWebElement GetFirstRow()
        {
            var scriptStr =
                "try {" +
                "return document.querySelectorAll('[class^=\"viewportRow\"]')[0];" +
                "} catch(e) { return null; }";

            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)Driver;
            var element = (IWebElement)javaScriptExecutor.ExecuteScript(scriptStr);

            return element;
        }
    }
}

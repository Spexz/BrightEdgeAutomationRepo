using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
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

        /// <summary>
        /// Login the user in chrome
        /// </summary>
        /// <param name="user">The <see cref="User"/> with email and password</param>
        public static void LoginUser(User user)
        {
            if (Driver.ElementExist(By.Name("data[User][login]")))
            {
                // Try to login
                var emailField = Driver.FindElement(By.Name("data[User][login]"), 15);
                emailField.SendKeys(user.Email);

                var passwordField = Driver.FindElement(By.Name("data[User][password]"), 15);
                passwordField.SendKeys(user.Password);

                var loginButton = Driver.FindElement(By.Id("login_submit"), 15);
                loginButton.Click();

                Thread.Sleep(1000);

                Driver.Navigate().GoToUrl(BaseURL);

            }
        }

        /// <summary>
        /// Clean up Bright Edge
        /// </summary>
        public static void DeleteQueries()
        {
            if (Driver.ElementExist(By.XPath("//div/a[button/span[contains(., 'All Queries')]]")))
            {
                var allQueries = Driver.FindElement(By.XPath("//div/a[button/span[contains(., 'All Queries')]]"), 15);
                Click(allQueries);
            }

            // (//div[input[contains(@class, 'toggle')]])[2]
            if (Driver.ElementExist(By.XPath("(//div[input[contains(@class, 'toggle')]])[2]")))
            {
                var checkTopRow = Driver.FindElement(By.XPath("(//div[input[contains(@class, 'toggle')]])[2]"), 15);
                Click(checkTopRow);

                if (Driver.ElementExist(By.XPath("//button[span[contains(., 'Delete Queries')]]")))
                {
                    var deleteQueries = Driver.FindElement(By.XPath("//button[span[contains(., 'Delete Queries')]]"), 15);
                    Click(deleteQueries);
                }
            }
        }

        public static void DeleteQueries2(string keywords, string location)
        {
            
            if (Driver.ElementExist(By.XPath("//div/a[button/span[contains(., 'All Queries')]]")))
            {
                var allQueries = Driver.FindElement(By.XPath("//div/a[button/span[contains(., 'All Queries')]]"), 15);
                Click(allQueries);
            }

            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(15));
            Func<IWebDriver, IWebElement> waitForElement = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
            {
                try
                {
                    var el = GetFirstRowToggleWithParam(keywords, location);

                    if (el != null)
                        return el;

                    return null;
                }
                catch (Exception e)
                {
                    return null;
                }
            });

            IWebElement targetElement = wait.Until(waitForElement);

            if(targetElement != null)
            {
                Click(targetElement);

                // Delete the row
                if (Driver.ElementExist(By.XPath("//button[span[contains(., 'Delete Queries')]]")))
                {
                    var deleteQueries = Driver.FindElement(By.XPath("//button[span[contains(., 'Delete Queries')]]"), 15);
                    Click(deleteQueries);
                }

                Thread.Sleep(5000);
            }
        }

        /// <summary>
        /// Remove country from the country selector field 
        /// </summary>
        /// <param name="location">The country to be removed</param>
        public static void RemoveLocation(string location)
        {
            
            var countryXPath = $"//div[contains(@data-testid, 'locations')]//span[contains" +
                $"(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), " +
                $"'{location.Trim().ToLower()}')]/following-sibling::span";

            if (Driver.ElementExist(By.XPath(countryXPath)))
            {
                // Add location click
                var locationRemove = Driver.FindElement(By.XPath(countryXPath), 15);
                Click(locationRemove);
            }
        }

        /// <summary>
        /// Run the main Bright Edge keyword stats lookup process including downloading the csv stats
        /// </summary>
        /// <param name="keywords">The keywords to process</param>
        /// <param name="location">The country</param>
        public static void RunProcess(string keywords, string location)
        {
            RetryUntilSuccessOrTimeout(() => {
                try
                {
                    ActionProcess(keywords, location);
                    return true;
                }
                catch (Exception e)
                {
                    Driver.Navigate().GoToUrl(BaseURL);
                    Thread.Sleep(5000);

                    return false;
                }
            }, TimeSpan.FromMinutes(5));

        }


        public static void ActionProcess(string keywords, string location)
        {
            var keywordField = Driver.FindElement(By.Id("keywords"), 15);
            keywordField.Clear();
            System.Windows.Clipboard.SetText(keywords);
            keywordField.SendKeys(OpenQA.Selenium.Keys.Control + "v");

            //check if locations already added
            //div[contains(@data-testid, 'locations')]//span[contains(text(), 'Thailand')]
            var countryXPath = $"//div[contains(@data-testid, 'locations')]//span[contains(translate(text()," +
                $"'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{location.ToLower()}')]";

            if (!Driver.ElementExist(By.XPath(countryXPath)))
            {
                // Add location click
                var addLocation = Driver.FindElement(By.XPath("//div[contains(@data-testid, 'locations')]"), 15);
                Click(addLocation);

                // Input location
                var inputLocation = Driver.FindElement(By.XPath("//input[contains(@name, 'searchTerm')]"), 15);
                //Click(inputLocation);
                inputLocation.Clear();
                System.Windows.Clipboard.SetText(location);
                inputLocation.SendKeys(OpenQA.Selenium.Keys.Control + "v");

                // Selection country checkbox
                var countryCheckXpath = $"//div[input[translate(@id," +
                    $"'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')= '{location.ToLower()}']]";
                //var checkLocation = Driver.FindElement(By.XPath($"//div[input[@id= '{location}']]"), 15);
                var checkLocation = Driver.FindElement(By.XPath(countryCheckXpath), 15);
                Click(checkLocation);

                // collapse add location
                //xpath: //div[contains(@data-testid, 'locations')]
                var collapseLocation = Driver.FindElement(By.XPath("//div[contains(@data-testid, 'locations')]"), 15);
                Click(collapseLocation);
            }



            //var initialDetailsUrl = GetDetailsUrl();

            // Get volume button
            var getVolume = Driver.FindElement(By.XPath("//button[contains(@data-testid, 'queryButton')]"), 15);
            Click(getVolume);

            /*RetryUntilSuccessOrTimeout(() => {
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

            }, TimeSpan.FromSeconds(120));

            // Click the View Details button
            // //div[contains(@class, 'viewportRow')]//button
            var viewDetailsButton = Driver.FindElement(By.XPath("//div[contains(@class, 'viewportRow')]//button"), 15);
            Click(viewDetailsButton);*/

            IWebElement viewResultsButton = null;
            RetryUntilSuccessOrTimeout(() => {
                try
                {
                    viewResultsButton = GetFirstRowLinkWithParam(keywords, location);

                    if (viewResultsButton != null)
                        return true;

                    return false;
                }
                catch (Exception e)
                {
                    return false;
                }

            }, TimeSpan.FromSeconds(120));
            Click(viewResultsButton);


            // Now download the csv
            var downloadButton = Driver.FindElement(By.XPath("//div[contains(@class, 'popover__target') and span/button[span[contains(text(), 'd')]]]"), 15);
            Click(downloadButton);

            Thread.Sleep(2000);

            //DeleteQueries();
            DeleteQueries2(keywords, location);
        }


        private static bool Click(IWebElement element)
        {
            var result = RetryUntilSuccessOrTimeout(() => {

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

            return result;
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

        private static IWebElement GetFirstRowWithParam(string keywords, string location)
        {
            var scriptStr =
            @"
                function getKeywordRow(mKeywords, mLocation)
                {
                    for(var i = 0; i < 3; i++)
                    {
                        try{
                            mKeywords = mKeywords.replace(/(?:\r\n|\r|\n)/g, ' ');
                            var row = document.querySelectorAll('[class^=\""viewportRow""\]')[i];
                            var keywordCell = row.querySelectorAll('[class^=\""cell""\]')[2];
                            var locationCell = row.querySelectorAll('[class^=\""cell""\]')[4];
                            var keywords = keywordCell.querySelector('span').innerText;

                            var location = locationCell.innerText.replace('United States', '').replace(/[^\w\s]/gi, '').trim();
                            keywords = keywords.replace(/[^\w\s]/gi, '');

                            //console.log(mKeywords);
                            //console.log(keywords);

                            if (mKeywords.toLowerCase().startsWith(keywords.toLowerCase()) &&
                                mLocation.toLowerCase().includes(location.toLowerCase()))
                            {
                                return rows;
                            }
                        } catch(e){}

                    }
                    return null;
                }

                var result = getKeywordRow(arguments[0], arguments[1]);
                return result;
            ";


            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)Driver;
            var element = (IWebElement)javaScriptExecutor.ExecuteScript(scriptStr, keywords.Substring(0,100), location);

            return element;
        }

        public static IWebElement GetFirstRowLinkWithParam(string keywords, string location)
        {
            var scriptStr =
            @"
                function getKeywordRowLink(mKeywords, mLocation)
                {
                    for(var i = 0; i < 10; i++)
                    {
                        try
                        {
                            mKeywords = mKeywords.replace(/(?:\r\n|\r|\n)/g, ' ');
                            var row = document.querySelectorAll('[class^=\""viewportRow""\]')[i];
                            var keywordCell = row.querySelectorAll('[class^=\""cell""\]')[2];
                            var locationCell = row.querySelectorAll('[class^=\""cell""\]')[4];
                            var keywords = keywordCell.querySelector('span').innerText;

                            var location = locationCell.innerText.replace('United States', '').replace(/[^\w\s]/gi, '').trim();
                            keywords = keywords.replace(/[^\w\s]/gi, '');

                            console.log(mKeywords);
                            console.log(keywords);

                            if (mKeywords.toLowerCase().startsWith(keywords.toLowerCase()) &&
                                mLocation.toLowerCase().includes(location.toLowerCase()))
                            {
                                return row.querySelector('button');
                            }
                        } catch(e){}

                    }
                    return null;
                }

                var result = getKeywordRowLink(arguments[0], arguments[1]);
                return result;
            ";


            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)Driver;
            var element = (IWebElement)javaScriptExecutor.ExecuteScript(scriptStr, keywords.Substring(0, 100), location);

            return element;
        }

        private static IWebElement GetFirstRowToggleWithParam(string keywords, string location)
        {
            var scriptStr =
            @"
                function getKeywordRowToggle(mKeywords, mLocation)
                {
                    for(var i = 0; i < 3; i++)
                    {
                        try{
                            mKeywords = mKeywords.replace(/(?:\r\n|\r|\n)/g, ' ');
                            var row = document.querySelectorAll('[class^=\""viewportRow""\]')[i];
                            var keywordCell = row.querySelectorAll('[class^=\""cell""\]')[2];
                            var locationCell = row.querySelectorAll('[class^=\""cell""\]')[4];
                            var keywords = keywordCell.querySelector('span').innerText;

                            var location = locationCell.innerText.replace('United States', '').replace(/[^\w\s]/gi, '').trim();
                            keywords = keywords.replace(/[^\w\s]/gi, '');

                            //console.log(mKeywords);
                            //console.log(keywords);

                            if (mKeywords.toLowerCase().startsWith(keywords.toLowerCase()) &&
                                mLocation.toLowerCase().includes(location.toLowerCase()))
                            {
                                return row.querySelector('div[class^=\""toggle""\]');
                            }
                        } catch(e){}

                    }
                    return null;
                }

                var result = getKeywordRowToggle(arguments[0], arguments[1]);
                return result;
            ";


            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)Driver;
            var element = (IWebElement)javaScriptExecutor.ExecuteScript(scriptStr, keywords.Substring(0, 100), location);

            return element;
        }
    }
}

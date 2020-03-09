using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace BrightEdgeAutomationTool
{
    public static class WebDriverExtensions
    {
        public static IWebElement FindElement(this IWebDriver driver, By by, int timeoutInSeconds)
        {
            Thread.Sleep(1000);
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));  //15
            Func<IWebDriver, IWebElement> waitForElement = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
            {
                try
                {
                    var el = driver.GetElementJS(by);

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
            return targetElement;
        }

        public static bool ElementExist(this IWebDriver driver, By by)
        {

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1.5));
                Func<IWebDriver, bool> waitForElement = new Func<IWebDriver, bool>((IWebDriver Web) =>
                {
                    var element = driver.GetElementJS(by);

                    if (element != null)
                        return true;
                    else
                        return false;

                });
                var result = wait.Until(waitForElement);
                return result;
            }
            catch(Exception e)
            {
                return false;
            }
            
        }

        public static IWebElement GetElementJS(this IWebDriver driver, By by)
        {

            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)driver;

            javaScriptExecutor.ExecuteScript("window.focus();");

            string window = driver.CurrentWindowHandle;
            driver.SwitchTo().Window(window);

            var test = by.ToString();
            Char[] spearator = { ':'};

            IWebElement element = null;
            var method = by.ToString().Split(spearator)[0].Trim();

            switch (method)
            {
                case "By.XPath":
                    element = (IWebElement)javaScriptExecutor.ExecuteScript(
                        //"console.log(arguments[0]);" +
                        "function getElementByXpath(path) {" +
                            //"console.log(document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue);" +
                            "return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; }" +
                            "" +
                        "return getElementByXpath(arguments[0]);", by.ToString().Replace("By.XPath: ", ""));
                    break;
                case "By.Name":
                    element = (IWebElement)javaScriptExecutor.ExecuteScript(
                        "var elements = document.getElementsByName(arguments[0]);"+
                        "if (elements.length > 0) return elements[0];" + 
                        "else return null;", by.ToString().Replace("By.Name: ", ""));
                    break;
                case "By.Id":
                    element = (IWebElement)javaScriptExecutor.ExecuteScript(
                        "var element = document.getElementById(arguments[0]);" +
                        "return element;", by.ToString().Replace("By.Id: ", ""));
                    break;
            }


            

            return element;

        }

        public static void JSClick(this IWebDriver driver, IWebElement element)
        {
            IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)driver;

            javaScriptExecutor.ExecuteScript("try {arguments[0].click();return true;}catch(e){return false;}", element);
        }

    }




}

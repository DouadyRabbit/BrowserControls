using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;

using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;


namespace BrowserControls
{
   
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("BrowserControl.EdgeControl")]
    public class EdgeControl
    {
        public string binaryLocation { get; set; }
        public string driverDir { get; set; }
        private IWebDriver webDriver { get; set; }

        public string errMsg = "";

        public EdgeControl()
        {
            var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            driverDir = outPutDirectory + @"\BrowserDrivers";
            binaryLocation = String.IsNullOrEmpty(binaryLocation) ? @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" : binaryLocation;
            InitDriver();
        }

        /// <summary>
        /// Initialize the Edge driver, same code works for other browsers
        /// </summary>
        /// <returns></returns>
        private int InitDriver()
        {
            int rtnVal = 0;
            var driverService = EdgeDriverService.CreateDefaultService(driverDir);
            driverService.HideCommandPromptWindow = true;
        
            EdgeOptions edgeOptions = new EdgeOptions();
         
            edgeOptions.BinaryLocation = binaryLocation;
            edgeOptions.AddArguments("--noerrdialogs");

            edgeOptions.AddArgument("no-sandbox");
            edgeOptions.AddArgument("start-minimized");

            edgeOptions.AddAdditionalOption("useAutomationExtension", false);
            edgeOptions.AddAdditionalOption("edgeDriverDirectory", driverDir);
          
            try
            {
                webDriver = new EdgeDriver(driverService, edgeOptions);
                webDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            }
            catch (System.Exception ex)
            {

                errMsg = ex.Message;
                rtnVal = -1;
            }

            return rtnVal;
        }

        /// <summary>
        /// Navigage2
        /// </summary>
        /// <param name="url"></param>
        public void navigate2(string url)
        {
            webDriver.Navigate().GoToUrl(url);
            WaitForLoad(webDriver);
        }


        /// <summary>
        /// Click and element  given the XPath, 
        /// </summary>
        /// <param name="xPathAddress"></param>
        [HandleProcessCorruptedStateExceptions]
        public int ClickByXPath(string xPathAddress)
        {
            int rtnVal = 0;

            WebDriverWait wait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(30));
            //wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(WebDriverTimeoutException));

            IJavaScriptExecutor js = (IJavaScriptExecutor)webDriver;

            //Down and up.
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
            js.ExecuteScript("window.scrollTo(0, 0)");

            //This code fails to trap exceptions since the exceptions might occur within the Wait function, see commented code at
            //Bottom of this document for possible solutions. This is commonly asked on StackOverflow.
            try
            {
                //var element = webDriver.FindElement(By.XPath(xPathAddress));
                IWebElement element = wait.Until(driver => driver.FindElement(By.XPath(xPathAddress)));

                //For zoom profile only, bring page upfront to make element clickable
                js.ExecuteScript("alert()");
                webDriver.SwitchTo().Alert().Accept();

                //Click element
                element.Click();
            }
            catch (NoSuchElementException ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }
            catch (Exception ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }

            return rtnVal;
        }


        /// <summary>
        /// setValueByID
        /// </summary>
        /// <param name="objectID"></param>
        /// <returns></returns>
        public int setValueByID(string objectID, string objectValue)
        {
            int rtnVal = 0;
            WebDriverWait wait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(30));
            //wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(WebDriverTimeoutException));

            //This code fails to trap exceptions since the exceptions might occur within the Wait function, see commented code at
            //Bottom of this document for possible solutions. This is commonly asked on StackOverflow.
            try
            {
                IWebElement element = wait.Until(driver => driver.FindElement(By.Id(objectID)));
        
                element.SendKeys(objectValue);
            }
            catch (NoSuchElementException ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }
            catch (Exception ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }



            return rtnVal;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="objectID"></param>
        /// <param name="objectValue"></param>
        /// <returns></returns>
        public int setValueByName(string objectName, string objectValue)
        {
            int rtnVal = 0;
            WebDriverWait wait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(30));
            //wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(WebDriverTimeoutException));

            //This code fails to trap exceptions since the exceptions might occur within the Wait function, see commented code at
            //Bottom of this document for possible solutions. This is commonly asked on StackOverflow.
            try
            {
                IWebElement element = wait.Until(driver => driver.FindElement(By.Name(objectName)));

                element.SendKeys(objectValue);
            }
            catch (NoSuchElementException ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }
            catch (Exception ex)
            {
                rtnVal = -1;
                errMsg = ex.Message;
            }



            return rtnVal;
        }

        /// <summary>
        /// Read inner HTML for element in xPathAddress
        /// </summary>
        /// <param name="xPathAddress"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public string ReadByXPath(string xPathAddress)
        {
            string elementValue = "";

            try
            {
                var oElement = webDriver.FindElement(By.XPath(xPathAddress));
                elementValue = oElement.Text;
            }
            catch (Exception ex)
            {

                errMsg = ex.Message;
            }
           

            return elementValue;
        }


        /// <summary>
        /// Maximizes browser
        /// </summary>
        public void Maximize()
        {
            webDriver.Manage().Window.Maximize();
        }

        /// <summary>
        /// Minimzes browser
        /// </summary>
        public void Minimize()
        {
            webDriver.Manage().Window.Minimize();
        }


        /// <summary>
        /// Wait for DOM to be complete, this might fail on pages with loading scripts.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="timeoutSec"></param>
        public static void WaitForLoad(IWebDriver driver, int timeoutSec = 15)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString().Trim().ToLower() == "complete");
        }

        /// <summary>
        /// Quit
        /// </summary>
        public void Quit()
        {
            webDriver.Quit();
        }

        /*
        ~EdgeControl()
        {
            Quit();
        }
        */
    }
}








/*
   IWebElement element = wait.Until<IWebElement>(d =>
            {
                try
                {
                    return d.FindElement(By.XPath(xPathAddress));
                }
                catch (WebDriverTimeoutException ex)
                {
                    rtnVal = -1;
                    errMsg = ex.Message;
                    return null;
                }
                catch (NoSuchElementException ex)
                {
                    rtnVal = -1;
                    errMsg = ex.Message;
                    return null;
                }
                catch (Exception ex)
                {
                    rtnVal = -1;
                    errMsg = ex.Message;
                    return null;
                }
            });

            if (rtnVal == 0)
            {
                js.ExecuteScript("alert()");
                webDriver.SwitchTo().Alert().Accept();
                element.Click();

            }
 */





/*
           //Wait for element to be enabled
           var element = wait.Until<IWebElement>(driver =>
           {
               try
               {
                   var elementToBeEnabled = driver.FindElement(By.XPath(xPathAddress));
                   if (elementToBeEnabled.Enabled)
                   {
                       return elementToBeEnabled;
                   }
                   return null;
               }
               catch (StaleElementReferenceException)
               {
                   return null;
               }
               catch (NoSuchElementException)
               {
                   return null;
               }
               catch (Exception ex)
               {
                   errMsg = ex.Message;
                   return null;
               }
           });

           */
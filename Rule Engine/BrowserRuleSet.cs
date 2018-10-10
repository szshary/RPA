using OpenQA.Selenium;
//using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace RPA.Core
{
    sealed public class BrowserRuleSet : RuleSetDecorator
    {
        private IWebDriver _htmlDriver;
        private WebDriverWait _wait;

        private List<String> _frameNames;

        public BrowserRuleSet(RuleSet _ruleSet) : base(_ruleSet)
        {
            Environment.SetEnvironmentVariable("webdriver.chrome.driver", Directory.GetCurrentDirectory() + "\\chromedriver.exe");
            Environment.SetEnvironmentVariable("webdriver.edge.driver", Directory.GetCurrentDirectory() + "\\MicrosoftWebDriver.exe");

            _elementStartRules.Add("AcceptAlert", AcceptAlert);
            _elementStartRules.Add("ClickBrowser", ClickBrowser);
            _elementStartRules.Add("SendTextToBrowser", SendTextToBrowser);
            _elementStartRules.Add("RefuseAlert", RefuseAlert);
            _elementStartRules.Add("WaitBrowser", WaitBrowser);

            _elementStartRules.Add("CompareAnyElementOfClassforVisibility", CompareAnyElementOfClassforVisibility);
            _elementEndRules.Add("CompareAnyElementOfClassforVisibility", PopConditionalStack);
            _elementStartRules.Add("CompareValueWithIdContent", CompareValueWithIdContent);
            _elementEndRules.Add("CompareValueWithIdContent", PopConditionalStack);
            _elementStartRules.Add("CompareVariableWithIdContent", CompareVariableWithIdContent);
            _elementEndRules.Add("CompareVariableWithIdContent", PopConditionalStack);

            _elementStartRules.Add("BrowserSession", StartBrowserSession);
            _elementEndRules.Add("BrowserSession", EndBrowserSession);
        }

        ~BrowserRuleSet()  // destructor
        {
            if (_htmlDriver != null)
            {
                _htmlDriver.Quit();
            }
        }

        private void AcceptAlert(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            try
            {
                _htmlDriver.SwitchTo().Alert().Accept();
            }
            catch (NoAlertPresentException)
            {
            }
        }

        private void ClickBrowser(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            try
            {
                RefreshHtmlFrames(parameters);
                IWebElement targetElement = null;
                if (parameters["Id"] != null)
                {
                    targetElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(parameters["Id"])));
                }
                else if (parameters.ContainsKey("ClassName"))
                {
                    targetElement = _htmlDriver.FindElement(By.ClassName(parameters["ClassName"]));
                }
                if (targetElement != null)
                {
                    _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(targetElement));
                    try
                    {
                        targetElement.Click();
                    }
                    catch (InvalidOperationException ex)
                    {
                    }
                }
            }
            catch (WebDriverTimeoutException)
            {
            }
        }

        private void CompareAnyElementOfClassforVisibility(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            bool isEqual = false;
            RefreshHtmlFrames(parameters);
            foreach (IWebElement elm in _htmlDriver.FindElements(By.ClassName(parameters["Class"])))
            {
                if (elm.Displayed)
                {
                    isEqual = true;
                    break;
                }
            }
            engineState.ConditionalStack.Push(isEqual);
        }

        private void CompareValueWithIdContent(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Value") && parameters.ContainsKey("Id"))
            {
                RefreshHtmlFrames(parameters);
                engineState.ConditionalStack.Push(parameters["Value"].ToString().Equals(_htmlDriver.FindElement(By.Id(parameters["Id"])).Text.Trim()));
            }
        }

        private void CompareVariableWithIdContent(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("Id"))
            {
                RefreshHtmlFrames(parameters);
                engineState.ConditionalStack.Push(engineState.VariableCollection[parameters["Variable"]].ToString().Equals(_htmlDriver.FindElement(By.Id(parameters["Id"])).Text.Trim()));
            }
        }

        private void RefuseAlert(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            try
            {
                _htmlDriver.SwitchTo().Alert().Dismiss();
            }
            catch (NoAlertPresentException)
            {
            }
        }

        private void SendTextToBrowser(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            try
            {
                RefreshHtmlFrames(parameters);
                IWebElement targetElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(parameters["Id"])));
                if (parameters.ContainsKey("Id") && targetElement != null)
                {
                    if (targetElement.Displayed)
                    {
                        try
                        {
                            targetElement.Clear();
                        }
                        catch (InvalidElementStateException)
                        {
                        }
                    }
                    else
                    {
                        IJavaScriptExecutor executor = (IJavaScriptExecutor)_htmlDriver;
                        executor.ExecuteScript("arguments[0].click();", targetElement);
                    }
                    if (parameters.ContainsKey("Value"))
                    {
                        targetElement.SendKeys(parameters["Value"]);
                    }
                    else if (parameters.ContainsKey("Variable") && engineState.VariableCollection.ContainsKey(parameters["Variable"]))
                    {
                        targetElement.SendKeys(engineState.VariableCollection[parameters["Variable"]].ToString());
                    }
                }
            }
            catch (WebDriverTimeoutException)
            {
            }
        }

        private void WaitBrowser(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("ExpectedCondition"))
            {
                try
                {
                    switch (parameters["ExpectedCondition"])
                    {
                        case "ElementIsVisible":
                            if (parameters.ContainsKey("Id"))
                            {
                                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id(parameters["Id"])));
                            }
                            break;
                        case "InvisibilityOfElementLocated":
                            if (parameters.ContainsKey("Id"))
                            {
                                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.Id(parameters["Id"])));
                            }
                            break;
                    }
                }
                catch (WebDriverTimeoutException)
                {
                }
            }
        }

        private void RefreshHtmlFrames(Dictionary<String, String> parameters)
        {
            _frameNames.Clear();
            _htmlDriver.SwitchTo().ParentFrame();
            if (parameters.ContainsKey("Frame"))
            {
                String[] frameStr = parameters["Frame"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (String str in frameStr)
                {
                    foreach (IWebElement webElm in _htmlDriver.FindElements(By.TagName("iframe")))
                    {
                        _frameNames.Add(webElm.GetAttribute("Id"));
                    }
                    if (_frameNames.Contains(str))
                    {
                        _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Id(str)));
                    }
                }
            }
            else
            {
                foreach (IWebElement webElm in _htmlDriver.FindElements(By.TagName("iframe")))
                {
                    _frameNames.Add(webElm.GetAttribute("Id"));
                }
            }
        }

        private void StartBrowserSession(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (_htmlDriver == null)
            {
                Process[] procs = null;
                try
                {
                    procs = Process.GetProcessesByName("chromedriver");

                    if (procs != null && procs.Length > 0)
                    {
                        Process chromeDriverProc = procs[0];

                        if (!chromeDriverProc.HasExited)
                        {
                            chromeDriverProc.CloseMainWindow();
                            chromeDriverProc.Close();
                        }
                    }
                }
                finally
                {
                    if (procs != null && procs.Length > 0)
                    {
                        foreach (Process p in procs)
                        {
                            p.Dispose();
                        }
                    }
                }

                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("user-data-dir=C:\\Users\\BIM2456\\AppData\\Local\\Google\\Chrome\\User Data\\Default");
                _htmlDriver = new ChromeDriver(chromeOptions);

                //htmlDriver = new EdgeDriver();

                //FirefoxProfile firefoxProfile = new FirefoxProfile("C:\\Path\\to\\profile");
                //htmlDriver = new FirefoxDriver(firefoxProfile);

                //htmlDriver = new InternetExplorerDriver();

                _htmlDriver.Url = parameters["URL"];
                _wait = new WebDriverWait(_htmlDriver, TimeSpan.FromSeconds(20.00));
                _wait.PollingInterval = TimeSpan.FromSeconds(2.0);

                _frameNames = new List<string>();
            }
        }

        private void EndBrowserSession(RuleEngineState engineState)
        {
            if (_htmlDriver != null)
            {
                _htmlDriver.Close();
                _htmlDriver.Quit();
            }
        }
    }
}
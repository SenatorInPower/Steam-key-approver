using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.Diagnostics;

class KeyAutomation
{
    public static int FindFreePort()
    {
        for (int port = 49152; port <= 65535; port++)
        {
            if (IsPortFree(port))
                return port;
        }
        throw new Exception("Не удалось найти свободный порт.");
    }

    private static bool IsPortFree(int port)
    {
        try
        {
            var listener = new TcpListener(IPAddress.Loopback, port);
            listener.Start();
            listener.Stop();
            return true;
        }
        catch (SocketException)
        {
            return false;
        }
    }
    public static void StartChromeWithDebugging(int port)
    {
        Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", $"--remote-debugging-port={port}");
    }

    static void Main()
    {
        var keyFilePath = @"C:\Users\Moskovchenko\Desktop\keys.txt";
        var resultFilePath = @"C:\Users\Moskovchenko\Desktop\results.txt";

        ChromeOptions options = new ChromeOptions();
        options.DebuggerAddress = "localhost:9222";  // Подключение к уже открытому браузеру
        IWebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        string initialUrl = "https://partner.steamgames.com/querycdkey/";
        List<string> keys = new List<string>(File.ReadAllLines(keyFilePath));
        List<string> results = new List<string>();

        // Переключение на первую вкладку
        driver.SwitchTo().Window(driver.WindowHandles[0]);

        foreach (var key in keys)
        {
            try
            {
                driver.Navigate().GoToUrl(initialUrl);

                IWebElement inputElement = wait.Until(ExpectedConditions.ElementIsVisible(By.Name("cdkey")));
                inputElement.Clear();
                inputElement.SendKeys(key);

                IWebElement sendButton = driver.FindElement(By.Name("method"));
                sendButton.Click();

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("span[style='color: #e24044']")));
                string activationStatus = driver.FindElement(By.CssSelector("span[style='color: #e24044']")).Text;

                string gameTitle = driver.FindElement(By.CssSelector("a[href*='packagelanding']")).Text;
                string packageId = driver.FindElement(By.CssSelector("a[href*='packagelanding']")).GetAttribute("href");

                string timeStamp = driver.FindElement(By.XPath("//td[contains(text(),'GMT')]")).Text;

                results.Add($"{key}: {activationStatus}, {gameTitle}, {packageId}, {timeStamp}");
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine($"Element not found: {e.Message}");
                results.Add($"{key}: Element not found.");
            }
            catch (WebDriverException e)
            {
                Console.WriteLine($"WebDriver error: {e.Message}");
                results.Add($"{key}: WebDriver error.");
            }
            Thread.Sleep(1000); // Для визуального контроля, можно убрать в продакшен-версии
        }

        driver.Quit();
        File.WriteAllLines(resultFilePath, results);
    }
}

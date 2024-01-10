using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;

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
        Console.WriteLine("Начало работы скрипта...");

        var keyFilePath = @"C:\Users\Moskovchenko\Desktop\keys.txt";
        var resultFilePath = @"C:\Users\Moskovchenko\Desktop\results.txt";

        Console.WriteLine("Настройка параметров Chrome...");
        ChromeOptions options = new ChromeOptions();
        options.DebuggerAddress = "localhost:9222";  // Подключение к уже открытому браузеру
        IWebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        string initialUrl = "https://partner.steamgames.com/querycdkey/";
        List<string> keys = new List<string>(File.ReadAllLines(keyFilePath));
        List<string> results = new List<string>();

        Console.WriteLine("Переключение на первую вкладку браузера...");
        driver.SwitchTo().Window(driver.WindowHandles[0]);

        foreach (var key in keys)
        {
            try
            {
                Console.WriteLine($"Обработка ключа: {key}");
                driver.Navigate().GoToUrl(initialUrl);
                Console.WriteLine("Страница загружена.");

                IWebElement inputElement = wait.Until(ExpectedConditions.ElementIsVisible(By.Name("cdkey")));
                inputElement.Clear();
                inputElement.SendKeys(key);

                IWebElement sendButton = driver.FindElement(By.Name("method"));
                sendButton.Click();
                Console.WriteLine("Запрос отправлен.");

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("span[style='color: #e24044']")));
                string activationStatus = driver.FindElement(By.CssSelector("span[style='color: #e24044']")).Text;

                string gameTitle = driver.FindElement(By.CssSelector("a[href*='packagelanding']")).Text;
                string packageId = driver.FindElement(By.CssSelector("a[href*='packagelanding']")).GetAttribute("href");

                string timeStamp = driver.FindElement(By.XPath("//td[contains(text(),'GMT')]")).Text;

                string result = $"{key}: {activationStatus}, {gameTitle}, {packageId}, {timeStamp}";
                Console.WriteLine($"Результат: {result}");
                results.Add(result);
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
            Thread.Sleep(1000); // Для визуального контроля
        }

        Console.WriteLine("Завершение работы браузера...");
        driver.Quit();

        Console.WriteLine("Сохранение результатов...");
        File.WriteAllLines(resultFilePath, results);

        Console.WriteLine("Работа скрипта завершена.");
    }

}

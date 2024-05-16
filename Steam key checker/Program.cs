using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Reflection;


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
    public static void StartChromeWithDebugging(int port, string chromePath)
    {
        Process.Start(chromePath, $"--remote-debugging-port={port}");
    }

    static void Main()
    {
        int processedKeyCount = 0;

      

      

        // Установка автоматического сброса буфера, чтобы записи лога немедленно выводились в файл
        Trace.AutoFlush = true;

        Console.WriteLine("Начало работы скрипта...");

        // Получение пути к рабочему столу пользователя
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        // Путь к файлу лога
        string logFilePath = $@"{desktopPath}\SteamKey\Steam-key-approver\Steam key checker\bin\Debug\net6.0\log.txt";

        // Добавление слушателя, который будет записывать логи в указанный файл
        Trace.Listeners.Add(new TextWriterTraceListener(logFilePath));

        // Определение путей к файлам на рабочем столе пользователя
        string keyFilePath = Path.Combine(desktopPath, "keys.txt");
        string resultFilePath = Path.Combine(desktopPath, "results.xlsx");

        // Удаление файла результатов, если он уже существует
        if (File.Exists(resultFilePath))
        {
            File.Delete(resultFilePath);
            Console.WriteLine("Существующий файл результатов был удален.");
        }


        // Get the full path of the current executable
        string exePath = Assembly.GetExecutingAssembly().Location;

        // Get the directory of the executable
        string chromeDriverPath = Path.GetDirectoryName(exePath);
        chromeDriverPath = Path.Combine(chromeDriverPath, "chromedriver.exe");
        // Check if the ChromeDriver file exists in the executable directory
        if (File.Exists(chromeDriverPath))
        {
            //   Console.WriteLine("ChromeDriver found at: " + chromeDriverPath);
            // Rest of your code that uses ChromeDriver...
        }
        else
        {
            Console.WriteLine("ChromeDriver not found in the executable directory.");
            // Handle the case where ChromeDriver is not found...
        }
        // Формирование нового пути к исполняемому файлу Chrome на основе пути к рабочему столу
        string chromePath = Path.Combine(desktopPath, @"SteamKey\Steam-key-approver\Steam key checker\bin\Debug\net6.0\chrome-win64\chrome.exe");

        if (!File.Exists(chromePath))
        {
            Console.WriteLine("Не удалось найти исполняемый файл Chrome.");
            // Здесь может быть код для обработки случая, когда файл не найден
        }
        else
        {
            // Если файл найден, продолжить выполнение программы
            Console.WriteLine($"Chrome найден по пути: {chromePath}");
            // Здесь ваш код, который использует chromePath
        }


        //@"C:\Program Files\Google\Chrome\Application\chrome.exe";
        if (!File.Exists(chromePath))
        {
            chromePath = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
            if (!File.Exists(chromePath))
            {
                // Если файл не найден в стандартных путях, ищем в системе
                chromePath = FindChromeExecutablePath();
                if (chromePath == null)
                {
                    Console.WriteLine("Не удалось найти исполняемый файл Chrome.");
                    return; // Завершаем выполнение, если Chrome не найден
                }
            }
        }


        // Запуск Chrome с удаленной отладкой
        int debugPort = 9222; // Порт для отладки
        StartChromeWithDebugging(debugPort, chromePath);

        // Настройка и запуск WebDriver
        ChromeDriverService service = ChromeDriverService.CreateDefaultService(chromeDriverPath);
        service.LogPath = "chromedriver.log";
        service.EnableVerboseLogging = true;
        try
        {
            var chromeDriverVersionInfo = Process.Start(new ProcessStartInfo
            {
                FileName = chromeDriverPath /*Path.Combine(chromeDriverPath, "chromedriver.exe")*/,
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            });

            Console.WriteLine("Версия ChromeDriver: " + chromeDriverVersionInfo.StandardOutput.ReadToEnd());
        }
        catch (Exception ex)
        {
            Console.WriteLine("Не удалось получить версию ChromeDriver: " + ex.Message);
        }
        ChromeOptions options = new ChromeOptions();
        options.AddArgument("--disable-extensions");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--no-sandbox");
        //  options.AddArgument("--headless");

        options.DebuggerAddress = $"localhost:{debugPort}";


        IWebDriver driver = new ChromeDriver(service, options);

        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        string urlForAnalysis = "https://partner.steamgames.com/querycdkey/cdkey?cdkey=&method=Query";

        driver.Navigate().GoToUrl(urlForAnalysis);

        List<string> keys = new List<string>(File.ReadAllLines(keyFilePath));
        List<string> results = new List<string>();

        Console.WriteLine("Переключение на первую вкладку браузера...");
        try
        {
            Trace.WriteLine("Попытка переключения на первую вкладку браузера...");
            driver.SwitchTo().Window(driver.WindowHandles[0]);
            Trace.WriteLine("Успешно переключились на первую вкладку браузера.");
        }
        catch (Exception e)
        {
            Trace.WriteLine($"Ошибка при попытке переключения на окно: {e.Message}");
            // Здесь можно добавить дополнительную логику обработки ошибки
        }

        // Создаем новую книгу или загружаем существующую
        XLWorkbook workbook;
        if (File.Exists(resultFilePath))
        {
            // Загрузить существующую книгу
            workbook = new XLWorkbook(resultFilePath);
        }
        else
        {
            // Создать новую книгу и лист
            workbook = new XLWorkbook();
            workbook.Worksheets.Add("Results");
        }

        // Получение или создание листа "Results"
        IXLWorksheet worksheet = workbook.Worksheet("Results");
        // Если лист новый, добавляем заголовки
        if (worksheet.LastRowUsed() == null)
        {
            worksheet.Cell("A1").Value = "Key";
            worksheet.Cell("B1").Value = "Status";
            worksheet.Cell("C1").Value = "Time Activation";
            worksheet.Cell("D1").Value = "Package";
            worksheet.Cell("E1").Value = "Tag";
        }

        int currentRow = worksheet.LastRowUsed()?.RowNumber() + 1 ?? 2; // Если лист пуст, начинаем с 2 строки

        int lastProcessedKeyIndex = -1;
        try
        {

            foreach (var key in keys)
            {

                Console.WriteLine($"Обработка ключа: {key}");
                driver.Navigate().GoToUrl(urlForAnalysis);
                Console.WriteLine("Страница загружена.");

                IWebElement inputElement = wait.Until(ExpectedConditions.ElementIsVisible(By.Name("cdkey")));
                inputElement.Clear();
                inputElement.SendKeys(key);

                IWebElement sendButton = driver.FindElement(By.Name("method"));
                sendButton.Click();
                Console.WriteLine("Запрос отправлен.");

                // Ожидаем появление статуса активации и получаем его текст
                string activationStatus = wait.Until(d =>
                {
                    var element = d.FindElement(By.CssSelector("span[style='color: #e24044'], span[style='color: #67c1f5']"));
                    return element.Displayed && element.Text.Length > 0 ? element.Text : null;
                });

                // Получаем название игры и ID пакета из URL
                string gameTitle = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("a[href*='packagelanding']"))).Text;
                string packageId = driver.FindElement(By.CssSelector("a[href*='packagelanding']")).GetAttribute("href");

                // Получаем временную метку активации
                string timeStamp = driver.FindElement(By.XPath("//td[contains(text(),'GMT')]")).Text;

                // Формируем результат
                string result = $"{key}: {activationStatus}, {gameTitle}, {packageId}, {timeStamp}";
                Console.WriteLine($"Результат: {result}");
                results.Add(result);

                // Записываем результаты в строку таблицы
                worksheet.Cell(currentRow, 1).Value = key;
                worksheet.Cell(currentRow, 2).Value = activationStatus;
                worksheet.Cell(currentRow, 3).Value = timeStamp;
                worksheet.Cell(currentRow, 4).Value = gameTitle;
                worksheet.Cell(currentRow, 5).Value = packageId; // Или другой способ получения Tag

                currentRow++;

                processedKeyCount++;

                // Проверяем, достигли ли мы 100 обработанных ключей
                if (processedKeyCount >= 30)
                {
                    // Сохраняем таблицу в файл и очищаем результаты
                    workbook.SaveAs(resultFilePath);
                    results.Clear();
                    Console.WriteLine($"Таблица сохранена после {processedKeyCount} запросов");

                    // Обновляем файл с ключами
                    UpdateKeyFile(keyFilePath, lastProcessedKeyIndex + 1);

                    // Сбрасываем счетчик
                    processedKeyCount = 0;
                }
             

                //catch (NoSuchElementException e)
                //{
                //    Console.WriteLine($"Element not found: {e.Message}");
                //    results.Add($"{key}: Element not found.");
                //}
                //catch (WebDriverException e)
                //{
                //    Console.WriteLine($"WebDriver error: {e.Message}");
                //    results.Add($"{key}: WebDriver error.");
                //}
                // После успешной обработки ключа, обновляем lastProcessedKeyIndex
                lastProcessedKeyIndex++;
                // Thread.Sleep(200); // Для визуального контроля
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка: {e.Message}");
            // В блоке catch больше не нужно сохранять результаты и закрывать WebDriver
        }
        finally
        {
            // Сохраняем таблицу в файл и закрываем WebDriver в блоке finally
            if (workbook != null)
            {
                Console.WriteLine("Сохранение результатов...");
                workbook.SaveAs(resultFilePath);
                Console.WriteLine("Результаты сохранены в Excel файл.");
            }

            if (driver != null)
            {
                Console.WriteLine("Завершение работы браузера...");
                driver.Quit();
            }

            if (lastProcessedKeyIndex >= 0)
            {
                UpdateKeyFile(keyFilePath, lastProcessedKeyIndex);
            }
        }



    }

    private static string FindChromeExecutablePath()
    {
        // Места, где обычно установлен Chrome
        string[] possiblePaths = {
            @"C:\Program Files\Google\Chrome\Application\chrome.exe",
            @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            // Добавьте дополнительные пути, если требуется
        };


        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
                return path;
        }

        // Ищем Chrome в реестре
        string keyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe";
        using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath))
        {
            if (key != null)
            {
                var path = key.GetValue(null) as string; // Получаем (по умолчанию) значение
                if (File.Exists(path))
                    return path;
            }
        }

        // Chrome не найден
        return null;
    }

    static void UpdateKeyFile(string keyFilePath, int lastProcessedKeyIndex)
    {
        var allKeys = File.ReadAllLines(keyFilePath).ToList();
        if (lastProcessedKeyIndex > 0 && lastProcessedKeyIndex < allKeys.Count)
        {
            var remainingKeys = allKeys.Skip(lastProcessedKeyIndex + 1).ToList();
            File.WriteAllLines(keyFilePath, remainingKeys);
        }
    }
}
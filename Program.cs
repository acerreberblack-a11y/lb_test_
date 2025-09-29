using Interop.UIAutomationClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Input;

namespace LandocsRobot
{

    internal class Program
    {
        // P/Invoke для BM_CLICK
        private const int BM_CLICK = 0x00F5;
        [DllImport("user32.dll")] static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;

        private static readonly Dictionary<string, string> _configValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _organizationValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _ticketValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static string _logFilePath = string.Empty;
        private static LogLevel _currentLogLevel = LogLevel.Info;

        #region Подключение утилит и параметры для них
        enum LogLevel
        {
            Fatal = 1,
            Error = 2,
            Warning = 3,
            Info = 4,
            Debug = 5
        }

        // Импорт функций из user32.dll
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void mouse_event(int dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        // Импорт функции из kernel32.dll
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetConsoleWindow();

        // Константы
        private const int SW_MINIMIZE = 6; // Команда для минимизации окна

        [Flags]
        private enum MouseFlags
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            Absolute = 0x8000
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        #endregion

        static void Main(string[] args)
        {
            // Основная логика робота

            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string configPath = Path.Combine(currentDirectory, "parameters.xml");
            string logDirectory = InitializeLogging();

            // Устанавливаем путь к файлу лога
            _logFilePath = Path.Combine(logDirectory, $"{DateTime.Now:yyyy-MM-dd}.log");
            Log(LogLevel.Info, "🤖 Запуск робота LandocsRobot");

            try
            {
                // Загрузка конфигураций
                if (!LoadConfig(configPath) || !LoadConfigOrganization(GetConfigValue("PathToOrganization")))
                {
                    Log(LogLevel.Error, "Ошибка при загрузке конфигурации. Завершение работы робота.");
                    return;
                }

                // Очистка старых файлов лога
                CleanOldLogs(logDirectory, int.TryParse(GetConfigValue("LogRetentionDays"), out int days) ? days : 30);

                string inputFolderPath = GetConfigValue("InputFolderPath");
                if (!Directory.Exists(inputFolderPath))
                {
                    Log(LogLevel.Error, $"Путь к папке входящих файлов [{inputFolderPath}] не существует. Завершение работы робота.");
                    return;
                }

                string OutputFolderPath = GetConfigValue("OutputFolderPath");
                if (!Directory.Exists(OutputFolderPath))
                {
                    Log(LogLevel.Error, $"Путь к папке входящих файлов [{OutputFolderPath}] не существует. Завершение работы робота.");
                    return;
                }

                //Получение входных файлов 
                string[] ticketArrays = Directory.GetDirectories(inputFolderPath);
                int ticketCount = ticketArrays.Length;

                Log(LogLevel.Info, ticketCount > 0
                    ? $"Найдено {ticketCount} заяв(-ка) (-ок) для обработки."
                    : "Папка пуста. Заявок для обработки не найдено.");

                if (ticketCount == 0)
                {
                    return;
                }

                foreach (string ticket in ticketArrays)
                {
                    try
                    {
                        // Очистка переменной заявки
                        _ticketValues.Clear();
                        string numberTicket = Path.GetFileNameWithoutExtension(ticket).Trim();
                        _ticketValues["ticketFolderName"] = numberTicket.Replace("+", "");

                        Log(LogLevel.Info, $"Начинаю обработку заявки: {numberTicket}");

                        // Поиск и проверка файла заявки
                        string ticketJsonFile = GetFileSearchDirectory(ticket, "*.txt");
                        if (ticketJsonFile == null)
                        {
                            Log(LogLevel.Error, $"Файл заявки [SD<Номер Заявки>.txt] не найден в папке [{ticket}]. Пропускаю заявку.");
                            continue;
                        }

                        Log(LogLevel.Info, $"Файл заявки [{Path.GetFileName(ticketJsonFile)}] найден. Начинаю обработку.");

                        // Парсинг JSON файла
                        var resultParseJson = ParseJsonFile(ticketJsonFile);
                        Log(LogLevel.Info, $"Извлеченные данные: Номер заявки - [{resultParseJson.Title}], Тип - [{resultParseJson.FormType}], Организация - [{resultParseJson.OrgTitle}], ППУД - [{resultParseJson.ppudOrganization}]");

                        // Сохранение извлеченной информации
                        _ticketValues["ticketName"] = resultParseJson.Title;
                        _ticketValues["ticketOrg"] = resultParseJson.OrgTitle;
                        _ticketValues["ticketType"] = resultParseJson.FormType;
                        _ticketValues["ticketPpud"] = resultParseJson.ppudOrganization;

                        // Поиск папки ЭДО
                        string ticketEdoFolder = GetFoldersSearchDirectory(ticket, "ЭДО");
                        if (ticketEdoFolder == null)
                        {
                            Log(LogLevel.Warning, $"Папка [ЭДО] не найдена в [{ticket}]. Пропускаю заявку.");
                            continue;
                        }

                        string[] ticketEdoChildren = GetFilesAndFoldersFromDirectory(ticketEdoFolder);
                        if (ticketEdoChildren.Length == 0)
                        {
                            Log(LogLevel.Error, $"Папка [ЭДО] пуста. Пропускаю заявку.");
                            continue;
                        }

                        Log(LogLevel.Info, $"В папке [ЭДО] найдено {ticketEdoChildren.Length} элементов. Начинаю обработку файлов.");

                        // Создание и проверка структуры папок
                        if (!EnsureDirectoriesExist(ticketEdoFolder, "xlsx", "pdf", "zip", "error", "document"))
                        {
                            Log(LogLevel.Error, $"Ошибка при создании структуры папок в [{ticketEdoFolder}]. Пропускаю заявку.");
                            continue;
                        }

                        // Сортировка и перемещение файлов
                        var newFoldersEdoChildren = CreateFolderMoveFiles(ticketEdoFolder, ticketEdoChildren);
                        Log(LogLevel.Info, "Сортировка и перемещение файлов завершены.");

                        // Логирование содержимого папок
                        Log(LogLevel.Debug, $"xlsx: {GetFileshDirectory(newFoldersEdoChildren.XlsxFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"pdf: {GetFileshDirectory(newFoldersEdoChildren.PdfFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"zip: {GetFileshDirectory(newFoldersEdoChildren.ZipFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"error: {GetFileshDirectory(newFoldersEdoChildren.ErrorFolder).Length} элементов.");

                        // Обработка файлов Excel
                        string[] xlsxFiles = XlsxContainsPDF(newFoldersEdoChildren.XlsxFolder, newFoldersEdoChildren.PdfFolder);
                        Log(LogLevel.Info, $"{xlsxFiles.Length} файл(-а) (-ов) на конвертацию в PDF.");

                        if (xlsxFiles.Length > 0)
                        {
                            ConvertToPdf(xlsxFiles, newFoldersEdoChildren.PdfFolder);
                            Log(LogLevel.Info, "Конвертация Excel в PDF завершена.");
                        }

                        // Сохранение пути к PDF
                        _ticketValues["pathPdf"] = newFoldersEdoChildren.PdfFolder;

                        Log(LogLevel.Info, $"Обработка заявки [{numberTicket}] завершена успешно.");
                    }
                    catch (Exception ticketEx)
                    {
                        Log(LogLevel.Error, $"Ошибка при обработке заявки [{ticket}]: {ticketEx.Message}");
                        continue;
                    }

                    //Обработка landocs
                    //Получаем списко файлов pdf для обработки
                    // Получаем файлы из директории, указанной в GetTicketValue("pathPdf")
                    string[] arrayPdfFiles = GetFilesAndFoldersFromDirectory(GetTicketValue("pathPdf")).ToArray();
                    /* string[] arrayPdfFiles = GetFilesAndFoldersFromDirectory(GetTicketValue("pathPdf"))
                         .Where(filePdf => !Path.GetFileName(filePdf).StartsWith("(+)", StringComparison.Ordinal))
                         .ToArray();*/
                    #region Начать обработку Landocs


                    // Проверяем, что arrayPdfFiles не пуст
                    if (arrayPdfFiles == null || arrayPdfFiles.Length == 0)
                    {
                        Log(LogLevel.Info, "Файлы для обработки не найдены. Запускаю Landocs и перехожу во вкладку [Сообщения].");
                        try
                        {
                            OpenLandocsAndNavigateToMessages();
                        }
                        catch (Exception ex)
                        {
                            Log(LogLevel.Error, $"Не удалось выполнить переход во вкладку [Сообщения]: {ex.Message}");
                        }
                        continue;
                    }

                    // Проверяем, все ли файлы содержат "+"
                    bool allFilesHavePlus = arrayPdfFiles.All(file => file.Contains("+"));

                    // Получаем путь к папке из первого файла
                    string folderPath = Path.GetDirectoryName(ticket);
                    if (string.IsNullOrEmpty(folderPath))
                    {
                        throw new InvalidOperationException("Invalid folder path.");
                    }

                    if (allFilesHavePlus)
                    {
                        // Формируем путь к целевой папке Output
                        string outputPath = Path.Combine(OutputFolderPath, Path.GetFileName(ticket));

                        // Создаем папку Output, если она не существует
                        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                        // Перемещаем папку
                        Directory.Move(folderPath, outputPath);
                        Console.WriteLine($"Заявка перемещена в : {outputPath}");
                        continue;
                    }
                        foreach (string filePdf in arrayPdfFiles)
                    {
                        string sourceFile = Path.GetFileName(filePdf);
                        if (Path.GetFileName(filePdf).Replace(" ", "").StartsWith("(+)", StringComparison.Ordinal))
                        {
                            continue;
                        }
                        int index = 0;
                        var resultparseFileName = GetParseNameFile(Path.GetFileNameWithoutExtension(filePdf));
                        Log(LogLevel.Info, $"Начинаю работу по файлу: Индекс: [{index}], Файл: [{sourceFile}]. Всего файлов: [{arrayPdfFiles.Length}]");
                        //Получаем наименование контрагента
                        _ticketValues["CounterpartyName"] = resultparseFileName.CounterpartyName?.Trim() ?? string.Empty;
                        //Получаем номер документа
                        _ticketValues["FileNameNumber"] = resultparseFileName.Number?.Trim() ?? string.Empty;
                        //Получаем дату документа
                        _ticketValues["FileDate"] = resultparseFileName.FileDate?.Trim() ?? string.Empty;
                        //Получаем ИНН
                        _ticketValues["FileNameINN"] = resultparseFileName.INN?.Trim() ?? string.Empty;
                        //Получаем КПП документа
                        _ticketValues["FileNameKPP"] = resultparseFileName.KPP?.Trim() ?? string.Empty;
                        try
                        {
                            Log(LogLevel.Info, $"Запускаю Landocs.");

                            // Получение путей из конфигурации
                            string customFile = GetConfigValue("ConfigLandocsCustomFile");  // Путь к исходному файлу
                            string landocsProfileFolder = GetConfigValue("ConfigLandocsFolder");  // Папка назначения

                            #region Запуск LanDocs

                            IUIAutomationElement appElement = null;
                            IUIAutomationElement targetWindowCreateDoc = null;
                            IUIAutomationElement targetWindowCounterparty = null;
                            IUIAutomationElement targetWindowAgreement = null;
                            IUIAutomationElement targetWindowGetPdfFile = null;
                            IUIAutomationElement targetWindowSettingFile2 = null;
                            IUIAutomationElement targetWindowSettingFile2SelectedWinodw = null;
                            IUIAutomationElement targetWindowsFinalCoordination = null;

                            IUIAutomationElement targetElementAgreementTree = null;
                            IUIAutomationElement targetWindowSettingFile = null;
                            string landocsProcessName = string.Empty;

                            try
                            {
                                // Перемещение пользовательского профиля Landocs
                                MoveCustomProfileLandocs(customFile, landocsProfileFolder);
                                Log(LogLevel.Info, "Профиль Landocs успешно перемещен.");

                                // Путь к приложению Landocs
                                string appLandocsPath = GetConfigValue("AppLandocsPath");
                                landocsProcessName = Path.GetFileNameWithoutExtension(appLandocsPath);

                                // Запуск приложения и ожидание окна
                                Log(LogLevel.Info, $"Запускаю приложение Landocs по пути: {appLandocsPath}");
                                appElement = LaunchAndFindWindow(appLandocsPath, "_robin_landocs (Мой LanDocs) - Избранное - LanDocs", 300);

                                if (appElement == null)
                                {
                                    Log(LogLevel.Error, "Окно Landocs не найдено. Завершаю работу.");
                                    throw new Exception("Окно Landocs не найдено.");
                                }

                                Log(LogLevel.Info, "Приложение Landocs успешно запущено и окно найдено.");

                                // Задержка на обработку интерфейса
                                Thread.Sleep(5000);
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при запуске Landocs: {ex.Message}");
                                throw;  // Пробрасываем исключение дальше
                            }
                            #endregion

                            #region Поиск вкладки "Главная"

                            // Поиск и клик по элементу "Главная" в ТабМеню
                            string xpathSettingAccount1 = "Pane[3]/Tab/TabItem[1]";
                            Log(LogLevel.Info, "Начинаю поиск вкладки [Главная] в навигационном меню...");

                            try
                            {
                                var targetElement1 = FindElementByXPath(appElement, xpathSettingAccount1, 60);

                                if (targetElement1 != null)
                                {
                                    Log(LogLevel.Info, "Вкладка [Главная] найдена. Выполняю клик.");
                                    ClickElementWithMouse(targetElement1);


                                    Log(LogLevel.Info, "Клик по вкладке [Главная] успешно выполнен.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти вкладку [Главная] в навигационном меню.");
                                    throw new Exception("Элемент не найден - вкладка [Главная] в навигационном меню.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по вкладке [Главная]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск слева в меню элемента "Документы"

                            string xpathSettingDoc = "Pane[1]/Pane/Pane[1]/Pane/Pane/Button[2]";
                            Log(LogLevel.Info, "Начинаю поиск кнопки [Документы] в навигационном меню...");

                            try
                            {
                                var targetElementDoc = FindElementByXPath(appElement, xpathSettingDoc, 60);
                                if (targetElementDoc != null)
                                {
                                    Log(LogLevel.Info, $"Нашел ссылку [Документы] в левом навигационном меню");
                                    TryInvokeElement(targetElementDoc);
                                    Log(LogLevel.Info, "Клик по элементу [Документы] успешно выполнен.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти вкладку [Главная] в навигационном меню.");
                                    throw new Exception("Элемент не найден - элемент [Документы] в навигационном меню.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по элементу [Документы]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Клик по элементу "Документы"
                            try
                            {
                                Log(LogLevel.Info, "Нажимаем Ctrl+F для вызова окна поиска ППУД.");
                                EnsureEnglishKeyboardLayout();
                                SendKeys.SendWait("^{f}");
                                Thread.Sleep(3000);

                                // Попытка получить элемент, который сейчас в фокусе
                                var targetElementSearch = GetFocusedElement();

                                // Значение ППУД из данных заявки
                                string ppudValue = GetTicketValue("ticketPpud");

                                if (targetElementSearch != null)
                                {
                                    Log(LogLevel.Info, "Элемент окна поиска ППУД успешно найден.");

                                    // Попытка получить паттерн ValuePattern для элемента
                                    if (targetElementSearch.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        // Устанавливаем значение через ValuePattern
                                        valuePattern.SetValue(ppudValue);
                                        Log(LogLevel.Info, "Значение введено в окно поиска ППУД через ValuePattern.");
                                    }
                                    else
                                    {
                                        // Если ValuePattern недоступен, используем SendKeys
                                        SendKeys.SendWait(ppudValue);
                                        Log(LogLevel.Warning, "ValuePattern недоступен. Значение введено в окно поиска ППУД через SendKeys.");
                                    }
                                }
                                else
                                {
                                    // Если элемент не найден, бросаем исключение
                                    Log(LogLevel.Error, "Не удалось найти элемент окна поиска ППУД.");
                                    throw new Exception("Элемент окна поиска ППУД не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при вводе значения в окно поиска ППУД: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск ППУД
                            try
                            {
                                Log(LogLevel.Info, "Ищу кнопку [Вниз] в окне поиска ППУД.");

                                // XPath для кнопки поиска вниз
                                string xpathSettingDown = "Pane[1]/Pane/Pane[1]/Pane/Pane/Pane/Pane/Tree/Pane/Pane/Pane/Button[3]";

                                // Поиск элемента
                                var targetElementDown = FindElementByXPath(appElement, xpathSettingDown, 60);

                                if (targetElementDown != null)
                                {
                                    // Устанавливаем фокус на элемент
                                    targetElementDown.SetFocus();
                                    Log(LogLevel.Info, "Фокус успешно установлен на кнопку [Вниз].");

                                    // Даем интерфейсу время для обработки фокуса
                                    Thread.Sleep(2000);

                                    Log(LogLevel.Info, "Нажали кнопку [Вниз] в окне поиска ППУД.");
                                    TryInvokeElement(targetElementDown);
                                    Log(LogLevel.Info, "Нажали кнопку [Вниз] в окне поиска ППУД успешно.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Элемент кнопки [Вниз] в окне поиска ППУД не найден.");
                                    throw new Exception("Элемент кнопки [Вниз] в окне поиска ППУД не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по кнопке [Вниз]: {ex.Message}");
                                throw;
                            }
                            #endregion
                            Thread.Sleep(2000);
                            #region Поиск элемента ППУД в списке Документов
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск элемента ППУД в списке документов.");

                                // XPath для группы элементов ППУД
                                string xpathSettingItem = "Pane[1]/Pane/Pane[1]/Pane/Pane/Pane/Pane/Tree/Group";

                                // Поиск элемента группы
                                IUIAutomationElement targetElementItem = FindElementByXPath(appElement, xpathSettingItem, 60);

                                // Значение ППУД из данных заявки
                                string ppudElement = GetTicketValue("ticketPpud");

                                if (targetElementItem != null)
                                {
                                    Log(LogLevel.Info, $"Группа элементов найдена. Ищу ППУД с значением: [{ppudElement}].");

                                    // Получение всех дочерних элементов
                                    IUIAutomationElementArray children = targetElementItem.FindAll(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (children != null && children.Length > 0)
                                    {
                                        bool isFound = false;

                                        for (int i = 0; i < children.Length; i++)
                                        {
                                            IUIAutomationElement item = children.GetElement(i);

                                            // Получение текстового значения элемента
                                            string value = item.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId)?.ToString() ?? "Нет значения";

                                            if (value == ppudElement)
                                            {
                                                // Вызов действия для найденного элемента
                                                try
                                                {
                                                    TryInvokeElement(item);
                                                    Log(LogLevel.Info, $"ППУД [{ppudElement}] найден и успешно обработан.");
                                                    isFound = true;
                                                    break;
                                                }
                                                catch
                                                {
                                                    Log(LogLevel.Error, $"Не удалось выполнить действие для ППУД [{ppudElement}].");
                                                    throw new Exception($"Ошибка: Не удалось выполнить действие для ППУД [{ppudElement}].");
                                                }
                                            }
                                        }

                                        if (!isFound)
                                        {
                                            Log(LogLevel.Error, $"ППУД [{ppudElement}] не найден в списке.");
                                            throw new Exception($"Ошибка: ППУД [{ppudElement}] не найден в списке документов.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Список элементов ППУД пуст или недоступен.");
                                        throw new Exception("Ошибка: Список элементов ППУД пуст или недоступен.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Группа с элементами ППУД не найдена.");
                                    throw new Exception("Ошибка: Группа с элементами ППУД не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске элемента ППУД: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Нажимаем кнопку "Создать документ"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск кнопки для создания документа.");

                                // XPath для кнопки
                                string xpathCreateDocButton = "Pane[3]/Pane/Pane/ToolBar[1]/Button";

                                // Поиск кнопки
                                var targetElementCreateDocButton = FindElementByXPath(appElement, xpathCreateDocButton, 60);

                                if (targetElementCreateDocButton != null)
                                {
                                    // Получение имени кнопки
                                    string elementValue = targetElementCreateDocButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString() ?? "Неизвестная кнопка";

                                    Log(LogLevel.Info, $"Кнопка [{elementValue}] найдена. Устанавливаю фокус и выполняю действие.");

                                    ClickElementWithMouse(targetElementCreateDocButton);
                                    Log(LogLevel.Info, $"Успешно нажали на кнопку [{elementValue}].");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Кнопка для создания документа не найдена.");
                                    throw new Exception("Ошибка: Кнопка для создания документа не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при нажатии на кнопку создания документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Окно "Создать документ"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск окна создания документа.");

                                string findNameWindow = "Без имени - Документ LanDocs";
                                targetWindowCreateDoc = FindElementByName(appElement, findNameWindow, 300);

                                string elementValue = null;

                                // Проверяем, был ли найден элемент
                                if (targetWindowCreateDoc != null)
                                {
                                    // Получаем значение свойства Name
                                    elementValue = targetWindowCreateDoc.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();


                                    // Проверяем, соответствует ли свойство Name ожидаемому значению
                                    if (elementValue == findNameWindow)
                                    {
                                        Log(LogLevel.Info, $"Появилось окно создания документа: [{elementValue}].");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, $"Ожидалось окно с названием 'Без имени - Документ LanDocs', но найдено: [{elementValue ?? "Неизвестное имя"}].");
                                        throw new Exception($"Неверное окно: [{elementValue ?? "Неизвестное имя"}].");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Окно создания документа не найдено.");
                                    throw new Exception("Окно создания документа не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске окна создания документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Выпадающий список "Тип документа"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю процесс выбора типа документа.");

                                // XPath для комбобокса и кнопки
                                string xpathElementTypeDoc = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[14]/ComboBox";
                                string xpathButtonTypeDoc = "Button[1]";
                                string typeDocument = "ППУД. Исходящий электронный документ";

                                // Поиск комбобокса
                                var targetElementTypeDoc = FindElementByXPath(targetWindowCreateDoc, xpathElementTypeDoc, 60);

                                if (targetElementTypeDoc != null)
                                {
                                    // Поиск кнопки внутри комбобокса
                                    var targetElementTypeDocButton = FindElementByXPath(targetElementTypeDoc, xpathButtonTypeDoc, 60);

                                    if (targetElementTypeDocButton != null)
                                    {
                                        // Фокус и клик по кнопке комбобокса
                                        targetElementTypeDocButton.SetFocus();
                                        TryInvokeElement(targetElementTypeDocButton);
                                        Log(LogLevel.Info, "Открыли список выбора типа документа.");

                                        // Поиск элемента типа документа по имени
                                        var docV = FindElementByName(targetWindowCreateDoc, typeDocument, 60);
                                        if (docV != null)
                                        {
                                            TryInvokeElement(docV);
                                            Log(LogLevel.Info, $"Выбрали тип документа: [{typeDocument}].");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, $"Элемент с именем '[{typeDocument}]' не найден.");
                                            throw new Exception($"Элемент с именем '[{typeDocument}]' не найден.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Кнопка комбобокса для выбора типа документа не найдена.");
                                        throw new Exception("Кнопка комбобокса для выбора типа документа не найдена.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Комбобокс для выбора типа документа не найден.");
                                    throw new Exception("Комбобокс для выбора типа документа не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при выборе типа документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Выпадающий список "Вид документа"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск и выбор типа документа для второго типа.");

                                // XPath для второго типа документа
                                string xpathElementTypeDocSecond = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[16]/ComboBox";
                                string typeDocumentSecond = "ППУД ИСХ. Акт сверки по договору / договорам";

                                // Поиск второго элемента ComboBox
                                var targetElementTypeDocSecond = FindElementByXPath(targetWindowCreateDoc, xpathElementTypeDocSecond, 60);

                                // Проверка, найден ли элемент
                                if (targetElementTypeDocSecond != null)
                                {
                                    // Поиск кнопки внутри ComboBox
                                    var targetElementTypeDocButtonSecond = FindElementByXPath(targetElementTypeDocSecond, "Button[1]", 60);

                                    if (targetElementTypeDocButtonSecond != null)
                                    {
                                        targetElementTypeDocButtonSecond.SetFocus();
                                        TryInvokeElement(targetElementTypeDocButtonSecond);
                                        Log(LogLevel.Info, "Нажали на кнопку выбора типа документа.");

                                        // Поиск и выбор второго типа документа по имени
                                        var docVSecond = FindElementByName(targetWindowCreateDoc, typeDocumentSecond, 60);
                                        if (docVSecond != null)
                                        {
                                            TryInvokeElement(docVSecond);
                                            Log(LogLevel.Info, $"Выбрали тип документа: [{typeDocumentSecond}].");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, $"Элемент с именем '[{typeDocumentSecond}]' не найден.");
                                            throw new Exception($"Элемент с именем '[{typeDocumentSecond}]' не найден.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Не удалось найти кнопку внутри ComboBox для второго типа документа.");
                                        throw new Exception("Не удалось найти кнопку внутри ComboBox для второго типа документа.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти ComboBox для второго типа документа.");
                                    throw new Exception("Не удалось найти ComboBox для второго типа документа.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при выборе типа документа для второго типа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка открытия списка контрагентов
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск кнопки для открытия окна с контрагентами.");

                                // XPath для кнопки "Открыть окно с контрагентами"
                                string xpathCounterpartyDocButton = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[7]/Edit/Button[1]";
                                var targetElementCounterpartyDocButton = FindElementByXPath(targetWindowCreateDoc, xpathCounterpartyDocButton, 60);

                                // Проверка, найден ли элемент
                                if (targetElementCounterpartyDocButton != null)
                                {
                                    // Попытка взаимодействия с кнопкой
                                    ClickElementWithMouse(targetElementCounterpartyDocButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [Открыть окно с контрагентами].");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Кнопка [Открыть окно с контрагентами] не найдена.");
                                    throw new Exception("Кнопка [Открыть окно с контрагентами] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при нажатии на кнопку [Открыть окно с контрагентами]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск окна с контрагентами (+ожидание 5 минут и разворачивание)

                            try
                            {
                                const string WINDOW_NAME = "Выбор элемента";
                                const string XPATH_WINDOW = "Window[1]";
                                const int TOTAL_TIMEOUT_SECONDS = 300; // 5 минут
                                const int POLL_DELAY_MS = 2000;        // каждые 2 секунды

                                targetWindowCounterparty = null;

                                DateTime deadline = DateTime.Now.AddSeconds(TOTAL_TIMEOUT_SECONDS);
                                while (DateTime.Now < deadline)
                                {
                                    // 1) Ищем по имени (с коротким внутренним таймаутом)
                                    targetWindowCounterparty = FindElementByName(targetWindowCreateDoc, WINDOW_NAME, 2);

                                    // 2) Если не нашли — ищем по XPath (также короткий таймаут)
                                    if (targetWindowCounterparty == null)
                                    {
                                        Log(LogLevel.Debug, "Окно не найдено по имени. Пробуем найти его по XPath...");
                                        targetWindowCounterparty = FindElementByXPath(targetWindowCreateDoc, XPATH_WINDOW, 2);
                                    }

                                    // 3) Если нашли — логируем, даём фокус и разворачиваем
                                    if (targetWindowCounterparty != null)
                                    {
                                        Log(LogLevel.Info, "Появилось окно поиска контрагента [Выбор элемента]. Разворачиваю на весь экран...");
                                        try
                                        {
                                            // фокус
                                            try { targetWindowCounterparty.SetFocus(); } catch { }

                                            // попытка через WindowPattern
                                            try
                                            {
                                                object winPatternObj = targetWindowCounterparty.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId);
                                                if (winPatternObj != null)
                                                {
                                                    var winPattern = (IUIAutomationWindowPattern)winPatternObj;
                                                    // 3 = Maximized в UIA (WindowVisualState_Maximized), но используем enum если у вас он доступен
                                                    winPattern.SetWindowVisualState(WindowVisualState.WindowVisualState_Maximized);
                                                }
                                                else
                                                {
                                                    // если WindowPattern недоступен — через WinAPI
                                                    TryMaximizeByWinApi(targetWindowCounterparty);
                                                }
                                            }
                                            catch
                                            {
                                                // резерв — через WinAPI
                                                TryMaximizeByWinApi(targetWindowCounterparty);
                                            }
                                        }
                                        catch (Exception exMax)
                                        {
                                            Log(LogLevel.Warning, "Не удалось развернуть окно на весь экран: " + exMax.Message);
                                        }

                                        // успех
                                        break;
                                    }

                                    // 4) Ждём и повторяем
                                    Thread.Sleep(POLL_DELAY_MS);
                                }

                                // 5) По итогам ожидания — проверяем
                                if (targetWindowCounterparty == null)
                                    throw new Exception("Окно поиска контрагента [Выбор элемента] не найдено в течение 300 секунд.");
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, "Ошибка при поиске окна поиска контрагента: " + ex.Message);
                                throw;
                            }
                            #endregion

                            #region Ищем элемент ввода контрагента
                            try
                            {
                                string xpatElementCounterpartyInput = "Pane[1]/Pane/Table/Pane/Pane/Edit/Edit[1]";
                                var targetElementCounterpartyInput = FindElementByXPath(targetWindowCounterparty, xpatElementCounterpartyInput, 60);

                                string counterparty = GetTicketValue("FileNameINN");

                                if (targetElementCounterpartyInput != null)
                                {
                                    // Проверяем, поддерживает ли элемент ValuePattern
                                    var valuePattern = targetElementCounterpartyInput.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;

                                    if (valuePattern != null)
                                    {
                                        valuePattern.SetValue(counterparty);
                                        Log(LogLevel.Info, $"Значение [{counterparty}] успешно введено в поле поиска контрагента через ValuePattern.");
                                    }
                                    else
                                    {
                                        // Если ValuePattern не поддерживается, используем SendKeys
                                        targetElementCounterpartyInput.SetFocus();
                                        SendKeys.SendWait(counterparty);
                                        Log(LogLevel.Info, $"Значение [{counterparty}] введено в поле поиска контрагента с помощью SendKeys.");
                                    }
                                }
                                else
                                {
                                    throw new Exception($"Поле поиска контрагента не найдено. Значение [{counterparty}] не удалось ввести.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при вводе значения в поле поиска контрагента: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Ищем элемент кнопка "Поиск" контрагента
                            try
                            {
                                string xpathSearchCounterpartyButton = "Pane[1]/Pane/Table/Pane/Pane/Button[2]";
                                var targetElementSearchCounterpartyButton = FindElementByXPath(targetWindowCounterparty, xpathSearchCounterpartyButton, 60);

                                if (targetElementSearchCounterpartyButton != null)
                                {
                                    var elementValue = targetElementSearchCounterpartyButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();
                                    if (elementValue != null)
                                    {
                                        targetElementSearchCounterpartyButton.SetFocus();
                                        TryInvokeElement(targetElementSearchCounterpartyButton);
                                        Log(LogLevel.Info, $"Нажали на кнопку поиска контрагента[{elementValue}]");
                                    }
                                    else
                                    {
                                        // Если ValuePattern не поддерживается, используем SendKeys
                                        targetElementSearchCounterpartyButton.SetFocus();
                                        SendKeys.SendWait("{Enter}");
                                        Log(LogLevel.Info, $"Нажали на кнопку поиска контрагента с помощью SendKeys.");
                                    }

                                }
                                else
                                {
                                    throw new Exception($"Элемент кнопки поиcка контрагента не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске элемента [Поиск] или клика по нему: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск контрагента (надёжные проверки наличия скролла, без вечных циклов) — C# 7.3
                            try
                            {
                                var automation = new CUIAutomation();

                                // --- входные критерии ---
                                string innValue = GetTicketValue("FileNameINN") ?? string.Empty;
                                string kppValue = GetTicketValue("FileNameKPP") ?? string.Empty;
                                string counterpartyName = GetTicketValue("CounterpartyName") ?? string.Empty;

                                // --- 1) Находим список и панель данных ---
                                string xpathCounterpartyList = "Pane[1]/Pane/Table";
                                Log(LogLevel.Info, "Ищем 'Список контрагентов' и 'Панель данных'...");

                                IUIAutomationElement list = FindElementByXPath(targetWindowCounterparty, xpathCounterpartyList, 60);
                                if (list == null) throw new Exception("Список контрагентов не найден.");

                                IUIAutomationElement dataPanel = FindElementByName(list, "Панель данных", 60);
                                if (dataPanel == null) throw new Exception("Не найдена 'Панель данных' внутри списка.");

                                // --- 2) Ждём появления непустых строк ---
                                const int loadTimeoutMs = 120000; // 2 мин
                                const int loadPollMs = 500;
                                int waited = 0;
                                bool hasRows = false;

                                while (waited < loadTimeoutMs)
                                {
                                    var children = dataPanel.FindAll(TreeScope.TreeScope_Children, automation.CreateTrueCondition());
                                    if (children != null && children.Length > 0)
                                    {
                                        for (int i = 0; i < children.Length; i++)
                                        {
                                            var di = FindElementByXPath(children.GetElement(i), "dataitem", 2);
                                            IUIAutomationValuePattern vp = di != null ? di.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern : null;
                                            if (vp != null && !string.IsNullOrWhiteSpace(vp.CurrentValue)) { hasRows = true; break; }
                                        }
                                    }
                                    if (hasRows) break;
                                    Thread.Sleep(loadPollMs);
                                    waited += loadPollMs;
                                }
                                if (!hasRows) throw new Exception("Строки списка контрагентов не появились в отведённое время.");

                                // --- локальные функции ---
                                void CollectVisible(List<string> lines, HashSet<string> seenSet)
                                {
                                    var children = dataPanel.FindAll(TreeScope.TreeScope_Children, automation.CreateTrueCondition());
                                    if (children == null) return;

                                    for (int i = 0; i < children.Length; i++)
                                    {
                                        var row = children.GetElement(i);
                                        var di = FindElementByXPath(row, "dataitem", 1);
                                        IUIAutomationValuePattern vp = di != null ? di.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern : null;

                                        if (vp != null)
                                        {
                                            var value = (vp.CurrentValue ?? string.Empty).Trim();
                                            if (value.Length == 0) continue;
                                            if (seenSet.Add(value)) lines.Add(value);
                                        }
                                    }
                                }

                                IUIAutomationElement FindVisibleTarget(string expectedLine)
                                {
                                    var children = dataPanel.FindAll(TreeScope.TreeScope_Children, automation.CreateTrueCondition());
                                    if (children == null) return null;

                                    for (int i = 0; i < children.Length; i++)
                                    {
                                        var row = children.GetElement(i);
                                        var di = FindElementByXPath(row, "dataitem", 1);
                                        IUIAutomationValuePattern vp = di != null ? di.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern : null;

                                        if (vp != null)
                                        {
                                            var line = (vp.CurrentValue ?? string.Empty).Trim();
                                            if (string.Equals(line, expectedLine, StringComparison.Ordinal))
                                                return di;
                                        }
                                    }
                                    return null;
                                }

                                bool SelectOrClick(IUIAutomationElement di)
                                {
                                    try
                                    {
                                        var si = di.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId) as IUIAutomationScrollItemPattern;
                                        if (si != null) { si.ScrollIntoView(); Thread.Sleep(120); }
                                    }
                                    catch { }

                                    try
                                    {
                                        var sel = di.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) as IUIAutomationSelectionItemPattern;
                                        if (sel != null)
                                        {
                                            try { di.SetFocus(); } catch { }
                                            sel.Select();
                                            Log(LogLevel.Info, "Контрагент выбран через SelectionItemPattern.Select().");
                                            return true;
                                        }
                                    }
                                    catch (Exception selEx)
                                    {
                                        Log(LogLevel.Warning, "Select() не удался: " + selEx.Message + ". Переходим к клику мышью.");
                                    }

                                    try
                                    {
                                        try { di.SetFocus(); } catch { }
                                        ClickElementWithMouse(di);
                                        Log(LogLevel.Info, "Контрагент выбран кликом мыши (fallback).");
                                        return true;
                                    }
                                    catch (Exception clickEx)
                                    {
                                        Log(LogLevel.Error, "Не удалось выбрать контрагента ни Select(), ни кликом: " + clickEx.Message);
                                        return false;
                                    }
                                }

                                // --- 3) Проверяем доступные механизмы прокрутки ---
                                var condScrollBar = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ScrollBarControlTypeId);
                                IUIAutomationElement scrollBar =
                                    list.FindFirst(TreeScope.TreeScope_Descendants, condScrollBar) ??
                                    targetWindowCounterparty.FindFirst(TreeScope.TreeScope_Descendants, condScrollBar);

                                // Возможные механизмы:
                                var rvOnBar = scrollBar != null ? scrollBar.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId) as IUIAutomationRangeValuePattern : null;
                                var spPanel = dataPanel.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;
                                var spList = list.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;

                                bool CanAnyScroll()
                                {
                                    // есть RangeValue на скроллбаре, или ScrollPattern на панели/списке, или хотя бы клавиши можно отправить (всегда true)
                                    return rvOnBar != null || spPanel != null || spList != null || true;
                                }

                                // Унифицированные «пробую вниз/вверх». Если НЕТ НИ ОДНОГО реального способа (кроме клавиш),
                                // тогда тут ставим явные "return false", чтобы внешние циклы могли корректно завершиться.
                                bool TryScrollDown()
                                {
                                    bool moved = false;
                                    // 1) RangeValue на ScrollBar
                                    if (rvOnBar != null)
                                    {
                                        double before = rvOnBar.CurrentValue;
                                        double min = rvOnBar.CurrentMinimum;
                                        double max = rvOnBar.CurrentMaximum;
                                        double step = rvOnBar.CurrentSmallChange > 0 ? rvOnBar.CurrentSmallChange : Math.Max((max - min) / 20.0, 1.0);
                                        double after = Math.Min(before + step, max);

                                        if (after > before)
                                        {
                                            try { rvOnBar.SetValue(after); Thread.Sleep(180); } catch { }
                                            double now = rvOnBar.CurrentValue;
                                            if (now > before + 1e-6) moved = true;
                                        }
                                    }

                                    // 2) ScrollPattern вниз
                                    if (!moved)
                                    {
                                        IUIAutomationScrollPattern sp = spPanel ?? spList;
                                        if (sp != null)
                                        {
                                            try { sp.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeIncrement); Thread.Sleep(120); moved = true; }
                                            catch
                                            {
                                                try { sp.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_SmallIncrement); Thread.Sleep(100); moved = true; } catch { }
                                            }
                                        }
                                    }

                                    // 3) Клавиши (как самый последний шанс)
                                    if (!moved)
                                    {
                                        // Если нет ни RV, ни ScrollPattern — считаем, что **реального** скролла нет.
                                        if (rvOnBar == null && spPanel == null && spList == null)
                                            return false; // <- ВАЖНО: говорим внешнему коду «не можем листать», чтобы не зациклиться.

                                        try { dataPanel.SetFocus(); } catch { try { list.SetFocus(); } catch { } }
                                        try { System.Windows.Forms.SendKeys.SendWait("{PGDN}"); Thread.Sleep(100); moved = true; } catch { }
                                        if (!moved)
                                            try { System.Windows.Forms.SendKeys.SendWait("{DOWN}{DOWN}{DOWN}"); Thread.Sleep(100); moved = true; } catch { }
                                    }
                                    return moved;
                                }

                                bool TryScrollUp()
                                {
                                    bool moved = false;
                                    // 1) RangeValue на ScrollBar
                                    if (rvOnBar != null)
                                    {
                                        double before = rvOnBar.CurrentValue;
                                        double min = rvOnBar.CurrentMinimum;
                                        double step = rvOnBar.CurrentSmallChange > 0 ? rvOnBar.CurrentSmallChange : Math.Max((rvOnBar.CurrentMaximum - min) / 20.0, 1.0);
                                        double after = Math.Max(before - step, min);

                                        if (after < before)
                                        {
                                            try { rvOnBar.SetValue(after); Thread.Sleep(180); } catch { }
                                            double now = rvOnBar.CurrentValue;
                                            if (now < before - 1e-6) moved = true;
                                        }
                                    }

                                    // 2) ScrollPattern вверх
                                    if (!moved)
                                    {
                                        IUIAutomationScrollPattern sp = spPanel ?? spList;
                                        if (sp != null)
                                        {
                                            try { sp.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeDecrement); Thread.Sleep(120); moved = true; }
                                            catch
                                            {
                                                try { sp.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_SmallDecrement); Thread.Sleep(100); moved = true; } catch { }
                                            }
                                        }
                                    }

                                    // 3) Клавиши
                                    if (!moved)
                                    {
                                        if (rvOnBar == null && spPanel == null && spList == null)
                                            return false; // <- тоже не можем листать.

                                        try { dataPanel.SetFocus(); } catch { try { list.SetFocus(); } catch { } }
                                        try { System.Windows.Forms.SendKeys.SendWait("{PGUP}"); Thread.Sleep(100); moved = true; } catch { }
                                        if (!moved)
                                            try { System.Windows.Forms.SendKeys.SendWait("{UP}{UP}{UP}"); Thread.Sleep(100); moved = true; } catch { }
                                    }
                                    return moved;
                                }

                                // --- 4) Сбор элементов (если скролл есть — идём вниз до упора; если нет — только видимые) ---
                                var allLines = new List<string>();
                                var seen = new HashSet<string>(StringComparer.Ordinal);

                                CollectVisible(allLines, seen);

                                const int collectTimeoutMs = 60000; // 1 мин максимум на «полную» прокрутку и сбор
                                int collectElapsed = 0;
                                int stableIters = 0;
                                int lastCount = allLines.Count;

                                if (rvOnBar != null || spPanel != null || spList != null)
                                {
                                    // попробовать к самому верху
                                    if (rvOnBar != null) { try { rvOnBar.SetValue(rvOnBar.CurrentMinimum); Thread.Sleep(150); } catch { } }
                                    else { try { dataPanel.SetFocus(); System.Windows.Forms.SendKeys.SendWait("{HOME}"); Thread.Sleep(120); } catch { } }

                                    // цикл сбора со скроллом
                                    while (collectElapsed < collectTimeoutMs)
                                    {
                                        bool moved = TryScrollDown();
                                        Thread.Sleep(120);
                                        collectElapsed += 120;

                                        CollectVisible(allLines, seen);

                                        if (allLines.Count > lastCount) { lastCount = allLines.Count; stableIters = 0; }
                                        else stableIters++;

                                        if (!moved && stableIters >= 3) break; // нечем листать/не движется и новых строк не появляется
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Info, "Скролл недоступен: работаем только с видимыми строками.");
                                }

                                if (allLines.Count == 0)
                                    throw new Exception("Не удалось собрать элементы из списка контрагентов.");

                                Log(LogLevel.Info, "Собрано строк контрагентов: " + allLines.Count);

                                // --- 5) Поиск ключа и строки ---
                                var counterpartyElements = new Dictionary<int, string[]>();
                                for (int i = 0; i < allLines.Count; i++)
                                    counterpartyElements[i] = allLines[i].Split(',').Select(v => v.Trim()).ToArray();

                                int? foundKey = FindCounterpartyKey(counterpartyElements, innValue, kppValue, counterpartyName);
                                if (!foundKey.HasValue)
                                {
                                    Log(LogLevel.Warning, "Контрагент не найден. Закрываю окно выбора контрагента.");
                                    SafeCloseWindow(targetWindowCounterparty, "Выбор контрагента");
                                    WaitWindowGoneByHandle(targetWindowCounterparty, 3000); // до 3 сек подождать
                                    throw new Exception("Не удалось найти контрагента по ИНН/КПП/Наименованию.");
                                }

                                string targetLine = allLines[foundKey.Value];
                                Log(LogLevel.Info, "Контрагент найден FindCounterpartyKey: индекс=" + foundKey.Value + ", строка=\"" + targetLine + "\".");

                                // --- 6) Доводим до видимости и выбираем ---
                                bool selected = false;

                                // Если прокрутка есть — пытаемся довести строку до видимой
                                if (rvOnBar != null || spPanel != null || spList != null)
                                {
                                    const int seekTimeoutMs = 90000; // 1.5 мин на доведение
                                    int seekElapsed = 0;

                                    // вниз
                                    while (seekElapsed < seekTimeoutMs && !selected)
                                    {
                                        var di = FindVisibleTarget(targetLine);
                                        if (di != null) { selected = SelectOrClick(di); break; }

                                        if (!TryScrollDown()) break; // <- если нечем/не движется — выходим
                                        Thread.Sleep(100);
                                        seekElapsed += 100;
                                    }

                                    // вверх
                                    if (!selected)
                                    {
                                        while (seekElapsed < seekTimeoutMs && !selected)
                                        {
                                            var di = FindVisibleTarget(targetLine);
                                            if (di != null) { selected = SelectOrClick(di); break; }

                                            if (!TryScrollUp()) break;
                                            Thread.Sleep(100);
                                            seekElapsed += 100;
                                        }
                                    }
                                }
                                else
                                {
                                    // скролл отсутствует: выбираем только если элемент уже виден
                                    var di = FindVisibleTarget(targetLine);
                                    if (di != null) selected = SelectOrClick(di);
                                    if (!selected)
                                    {
                                        Log(LogLevel.Warning, "Найденный контрагент не был выбран. Закрываю окно выбора контрагента.");
                                        SafeCloseWindow(targetWindowCounterparty, "Выбор контрагента");
                                        WaitWindowGoneByHandle(targetWindowCounterparty, 3000);
                                        throw new Exception("Не удалось довести до видимости и выбрать найденного контрагента.");
                                    }
                                }

                                if (!selected)
                                {
                                    Log(LogLevel.Warning, "Найденный контрагент не был выбран. Закрываю окно выбора контрагента.");
                                    SafeCloseWindow(targetWindowCounterparty, "Выбор контрагента");
                                    WaitWindowGoneByHandle(targetWindowCounterparty, 3000);
                                    throw new Exception("Не удалось довести до видимости и выбрать найденного контрагента.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, "Ошибка при поиске/выборе контрагента (скролл): " + ex.Message);
                                throw;
                            }
                            #endregion

                            #region Кнопка "Выбрать" в окне поиска контрагентов
                            try
                            {
                                string xpathCounterpartyOkButton = "Pane[2]/Button[1]";
                                Log(LogLevel.Info, "Начинаем поиск кнопки [Выбрать] в окне [Выбор элемента] со списком контрагентов...");

                                // Поиск кнопки [Выбрать]
                                var targetElementCounterpartyOkButton = FindElementByXPath(targetWindowCounterparty, xpathCounterpartyOkButton, 10);

                                if (targetElementCounterpartyOkButton != null)
                                {
                                    Log(LogLevel.Info, "Кнопка [Выбрать] найдена. Пытаемся нажать на кнопку...");

                                    // Установка фокуса на кнопку и попытка нажатия
                                    targetElementCounterpartyOkButton.SetFocus();
                                    TryInvokeElement(targetElementCounterpartyOkButton);

                                    Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор элемента] со списком контрагентов.");
                                }
                                else
                                {
                                    // Если кнопка не найдена
                                    throw new Exception("Кнопка [Выбрать] в окне [Выбор элемента] со списком контрагентов не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Выбрать]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [...] для открытия окна с договорами
                            try
                            {
                                string xpathAgreementButton = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[3]/Pane/Pane/Button[2]";
                                Log(LogLevel.Info, "Начинаем поиск кнопки [...] для выбора договора в окне [Создание документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementAgreementButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementButton, 10);

                                if (targetElementAgreementButton != null)
                                {
                                    Log(LogLevel.Info, "Кнопка открытия окна с  договорами найдена. Пытаемся нажать на кнопку...");

                                    // Установка фокуса на кнопку и попытка нажатия
                                    targetElementAgreementButton.SetFocus();
                                    ClickElementWithMouse(targetElementAgreementButton);

                                    Log(LogLevel.Info, "Нажали на кнопку открытия окна с договорами в окне [Создание документа].");
                                }
                                else
                                {
                                    // Если кнопка не найдена
                                    throw new Exception("Кнопка для выбора договора [...] в окне [Создание документа] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [...] для выбора договора: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск окна с договорами [Выбор документа]
                            try
                            {
                                // Поиск окна "Выбор документа"
                                targetWindowAgreement = FindElementByName(targetWindowCreateDoc, "Выбор документа", 60);

                                // Проверка, был ли найден элемент
                                if (targetWindowAgreement != null)
                                {
                                    Log(LogLevel.Info, "Окно поиска документа [Выбор документа] найдено.");
                                }
                                else
                                {
                                    Log(LogLevel.Warning, "Найденный контрагент не был выбран. Закрываю окно выбора документа.");
                                    SafeCloseWindow(targetWindowAgreement, "Выбор документа");
                                    WaitWindowGoneByHandle(targetWindowAgreement, 3000);
                                    throw new Exception("Ошибка: Окно поиска документа [Выбор документа] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске окна
                                Log(LogLevel.Error, $"Ошибка при поиске окна [Выбор документа]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск дерева [Журналы регистрации] в окне [Выбор документа]
                            try
                            {
                                // Поиск дерева элементов в списке [Журналы регистрации]
                                string xpathAgreementTree = "Pane/Pane/Pane[3]/Tree";
                                Log(LogLevel.Info, "Начинаем поиск дерева элементов списка [Журналы регистрации]...");

                                // Ищем элемент дерева
                                targetElementAgreementTree = FindElementByXPath(targetWindowAgreement, xpathAgreementTree, 60);

                                if (targetElementAgreementTree != null)
                                {
                                    Log(LogLevel.Info, "Элемент дерева [Журналы регистрации] найден.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Элемент дерева [Журналы регистрации] не найден. Работа робота завершена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске элемента
                                Log(LogLevel.Error, $"Ошибка при поиске элемента дерева [Журналы регистрации]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск скрола в дереве [Журнала регистрации]
                            try
                            {
                                Log(LogLevel.Info, "Элемент дерева [Журналы регистрации] найден. Пытаемся инициализировать скролл...");

                                // Поиск элемента скролла
                                var targetElementAgreemenScrollBar = FindElementByName(targetElementAgreementTree, "Vertical", 60);

                                if (targetElementAgreemenScrollBar != null)
                                {
                                    // Если элемент скролла найден
                                    Log(LogLevel.Info, "Элемент скролла [Vertical] найден! Работа робота продолжается.");
                                }
                                else
                                {
                                    // Если элемент скролла не найден
                                    throw new Exception("Ошибка: Элемент скролла [Vertical] не найден! Работа робота завершена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске скролла
                                Log(LogLevel.Error, $"Ошибка при поиске скролла [Vertical]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Проверка состояния элемента [Журналы регистраций]
                            try
                            {
                                // Поиск элемента [Журналы регистрации] в дереве
                                var targetElementAgreemenTreeItem = FindElementByName(targetElementAgreementTree, "Журналы регистрации", 60);

                                if (targetElementAgreemenTreeItem != null)
                                {
                                    Log(LogLevel.Info, "Элемент [Журналы регистраций] найден.");

                                    // Проверка, поддерживает ли элемент ExpandCollapsePattern
                                    if (targetElementAgreemenTreeItem.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) is IUIAutomationExpandCollapsePattern expandCollapsePattern)
                                    {
                                        var state = expandCollapsePattern.CurrentExpandCollapseState;

                                        switch (state)
                                        {
                                            case ExpandCollapseState.ExpandCollapseState_Collapsed:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] свернут. Раскрываем...");
                                                expandCollapsePattern.Expand(); // Раскрываем элемент
                                                Log(LogLevel.Info, "Элемент [Журналы регистраций] успешно раскрыт.");
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_Expanded:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] уже раскрыт.");
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_PartiallyExpanded:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] частично раскрыт. Раскрываем полностью...");
                                                expandCollapsePattern.Expand(); // Раскрываем элемент
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_LeafNode:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] является листовым узлом. Раскрытие не требуется.");
                                                break;

                                            default:
                                                Log(LogLevel.Warning, "Неизвестное состояние ExpandCollapseState.");
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент [Журналы регистраций] не поддерживает ExpandCollapsePattern.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Журналы регистраций] не найден.");
                                }

                                #region Поиск договора в дереве [Журналы регистраций] (Select() → fallback клик мышью)
                                try
                                {
                                    // Находим дочерние элементы
                                    IUIAutomationElementArray childrenAgreemen = targetElementAgreemenTreeItem.FindAll(
                                        TreeScope.TreeScope_Children,
                                        new CUIAutomation().CreateTrueCondition()
                                    );

                                    if (childrenAgreemen != null && childrenAgreemen.Length > 0)
                                    {
                                        bool isFound = false;
                                        int count = childrenAgreemen.Length;

                                        Log(LogLevel.Info, $"Количество журналов [{count}]");

                                        string agreementName = GetTicketValue("ticketPpud");
                                        var agreementNameSplit = agreementName.Split('.')[0]; // часть до первой точки
                                        var agreementNameFull = string.Concat(agreementNameSplit, ".", "Договоры").ToString();
                                        var agreementNameNormalize = agreementNameFull.Trim().ToLower().Replace(" ", "");

                                        // локальная функция: попытаться выбрать через Select(), иначе кликнуть мышью
                                        bool TrySelectOrClick(IUIAutomationElement childElement)
                                        {
                                            if (childElement == null) return false;

                                            string name = (childElement.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            Log(LogLevel.Debug, $"Фокус на элементе: [{name}]");

                                            if (agreementNameNormalize != name.Trim().ToLower().Replace(" ", ""))
                                                return false;

                                            Log(LogLevel.Info, $"Журнал [{agreementNameFull}] найден.");

                                            // Если поддерживается ScrollItemPattern — прокрутить в зону видимости
                                            if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId) is IUIAutomationScrollItemPattern scrollItemPattern)
                                            {
                                                try
                                                {
                                                    scrollItemPattern.ScrollIntoView();
                                                    Log(LogLevel.Debug, "Элемент журнала прокручен в область видимости.");
                                                    Thread.Sleep(300);
                                                }
                                                catch { /* не критично */ }
                                            }

                                            // Попытка выбора через SelectionItemPattern.Select()
                                            bool selected = false;
                                            try
                                            {
                                                if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern selectionItemPattern)
                                                {
                                                    try { childElement.SetFocus(); } catch { /* не критично */ }
                                                    selectionItemPattern.Select();
                                                    selected = true;
                                                    Log(LogLevel.Info, "Элемент журнала выбран через SelectionItemPattern.Select().");
                                                }
                                            }
                                            catch (Exception selEx)
                                            {
                                                Log(LogLevel.Warning, $"Select() не удался: {selEx.Message}. Падаем в клик мышью.");
                                            }

                                            // fallback: физический клик мышью
                                            if (!selected)
                                            {
                                                try
                                                {
                                                    try { childElement.SetFocus(); } catch { }
                                                    ClickElementWithMouse(childElement);
                                                    Log(LogLevel.Info, "Элемент журнала выбран кликом мыши (fallback).");
                                                    selected = true;
                                                }
                                                catch (Exception clickEx)
                                                {
                                                    Log(LogLevel.Error, $"Не удалось выбрать элемент ни Select(), ни кликом: {clickEx.Message}");
                                                    selected = false;
                                                }
                                            }

                                            return selected;
                                        }

                                        // 1) пробуем пройти текущие элементы
                                        for (int i = 0; i < count; i++)
                                        {
                                            var childElement = childrenAgreemen.GetElement(i);
                                            if (TrySelectOrClick(childElement))
                                            {
                                                isFound = true;
                                                break;
                                            }
                                        }

                                        // 2) если не нашли — скроллим вниз и ищем повторно
                                        if (!isFound)
                                        {
                                            Log(LogLevel.Debug, "Элемент не найден. Прокручиваем вниз.");
                                            var scrollPattern = targetElementAgreemenTreeItem.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;

                                            if (scrollPattern != null && scrollPattern.CurrentVerticallyScrollable != 0)
                                            {
                                                while (scrollPattern.CurrentVerticalScrollPercent < 100 && !isFound)
                                                {
                                                    scrollPattern.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeIncrement);
                                                    Log(LogLevel.Debug, "Прокручиваем вниз.");

                                                    // обновим список детей после прокрутки
                                                    childrenAgreemen = targetElementAgreemenTreeItem.FindAll(
                                                        TreeScope.TreeScope_Children,
                                                        new CUIAutomation().CreateTrueCondition()
                                                    );

                                                    for (int i = 0; i < childrenAgreemen.Length; i++)
                                                    {
                                                        var childElement = childrenAgreemen.GetElement(i);
                                                        if (TrySelectOrClick(childElement))
                                                        {
                                                            isFound = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (!isFound)
                                            throw new Exception("Журнал не найден или возникла ошибка при обработке элемента [Журналы регистраций].");
                                    }
                                    else
                                    {
                                        throw new Exception("Журнал не найден или возникла ошибка при обработке элемента [Журналы регистраций].");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log(LogLevel.Error, $"Ошибка при поиске журнала в дереве: {ex.Message}");
                                    throw;
                                }
                                #endregion


                                #region Поиск и нажатие кнопки "Выбрать"
                                try
                                {
                                    string xpathAgreementOkButton = "Pane/Pane/Pane[2]/Pane[3]/Button[1]";
                                    var targetElementAgreementOkButton = FindElementByXPath(targetWindowAgreement, xpathAgreementOkButton, 60);

                                    if (targetElementAgreementOkButton != null)
                                    {
                                        // Устанавливаем фокус на кнопку и нажимаем
                                        targetElementAgreementOkButton.SetFocus();
                                        TryInvokeElement(targetElementAgreementOkButton);
                                        Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор документа] со списком журналов.");
                                    }
                                    else
                                    {
                                        // Выбрасываем исключение, если элемент не найден
                                        throw new Exception("Кнопка [Выбрать] в окне [Выбор документа] со списком журналов не найдена.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Логируем ошибку и выбрасываем исключение дальше
                                    Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                    throw;
                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Проверяем, что договор проставлен
                            try
                            {
                                string xpathAgreementLabel = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[3]/Pane/Pane/Button[4]";
                                var targetElementAgreementLabelButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementLabel, 60);

                                if (targetElementAgreementLabelButton != null)
                                {
                                    // Получаем значение свойства Name
                                    string agreementLabelName = targetElementAgreementLabelButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string;

                                    // Проверяем, что значение не пустое
                                    if (!string.IsNullOrEmpty(agreementLabelName))
                                    {
                                        Log(LogLevel.Info, $"Договор проставлен успешно. Номер договора: {agreementLabelName}");
                                    }
                                    else
                                    {
                                        // Если значение пустое, выбрасываем исключение
                                        SafeCloseWindow(targetWindowAgreement, "Выбор документа");
                                        
                                        #region Выполняем повтоно поиск окна ВОПРОС, тк оно обновилось
                                        try
                                        {
                                            var automation = new CUIAutomation();
                                            var root = automation.GetRootElement();

                                            // --- параметры ожидания ---
                                            const int timeoutMs = 600000; // 10 минут
                                            const int pollMs = 500;       // 0.5 секунды
                                            int waited = 0;

                                            // --- ждём главное окно Landocs ---
                                            IUIAutomationElement mainWin = null;
                                            var condMainAutomationId = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                            var condMainTypeWindow = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                            var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                            while (waited < timeoutMs)
                                            {
                                                var candidates = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                                if (candidates != null && candidates.Length > 0)
                                                {
                                                    for (int i = 0; i < candidates.Length; i++)
                                                    {
                                                        var e = candidates.GetElement(i);
                                                        var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                        if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                                        {
                                                            mainWin = e;
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (mainWin != null) break;

                                                Thread.Sleep(pollMs);
                                                waited += pollMs;
                                            }

                                            if (mainWin == null)
                                                throw new Exception("Не найдено окно MainWindows (Name содержит 'Landocs') за 10 минут.");

                                            Log(LogLevel.Info, "Главное окно Landocs обнаружено.");

                                            // --- ждём DocCard внутри MainWindows ---
                                            IUIAutomationElement docCard = null;
                                            waited = 0;
                                            var condDocCardId = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                            var condDocTypeWindow = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                            var condDocTypePane = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                            var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                            var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                            while (waited < timeoutMs)
                                            {
                                                docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                                if (docCard != null)
                                                {
                                                    var docName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                    if (!string.IsNullOrEmpty(docName) && docName.StartsWith("Без имени", StringComparison.Ordinal))
                                                    {
                                                        Log(LogLevel.Info, $"DocCard найден: \"{docName}\".");
                                                        break;
                                                    }
                                                }
                                                Thread.Sleep(pollMs);
                                                waited += pollMs;
                                            }

                                            if (docCard == null)
                                                throw new Exception("DocCard не найден внутри окна Landocs за 10 минут.");

                                            var finalName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (!finalName.StartsWith("Без имени", StringComparison.Ordinal))
                                                throw new Exception($"Найден DocCard, но Name не начинается с \"Без имени\". Текущее Name: \"{finalName}\".");

                                            // --- ждём окно "Вопрос" внутри DocCard ---
                                            const int questionTimeoutMs = 120000; // 2 мин
                                            waited = 0;
                                            IUIAutomationElement questionWin = null;

                                            var condQuestionName = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_NamePropertyId, "Вопрос");
                                            var condTypeWindow = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                            var condTypePane = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                            var condType = automation.CreateOrCondition(condTypeWindow, condTypePane);
                                            var condQuestion = automation.CreateAndCondition(condQuestionName, condType);

                                            while (waited < questionTimeoutMs)
                                            {
                                                questionWin = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                                if (questionWin != null)
                                                {
                                                    bool isEnabled = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                    bool isOffscreen = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                                    if (isEnabled && !isOffscreen)
                                                        break;
                                                    questionWin = null;
                                                }

                                                Thread.Sleep(pollMs);
                                                waited += pollMs;
                                            }

                                            if (questionWin == null)
                                                throw new Exception("Внутри DocCard не найдено доступное окно с Name='Вопрос'.");

                                            try { questionWin.SetFocus(); } catch { }
                                            Log(LogLevel.Info, "Окно 'Вопрос' внутри DocCard найдено и активировано.");

                                            // --- ищем кнопку ОК внутри окна "Вопрос" ---
                                            IUIAutomationElement okBtn = null;
                                            waited = 0;
                                            string[] okNames = { "&НЕТ", "&Нет", "Нет", "нет" }; // оставил твои варианты и добавил частые

                                            var condOkType = automation.CreatePropertyCondition(
                                                UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ButtonControlTypeId);

                                            while (waited < questionTimeoutMs && okBtn == null)
                                            {
                                                foreach (var nm in okNames)
                                                {
                                                    var condOkName = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, nm);
                                                    var condOk = automation.CreateAndCondition(condOkName, condOkType);
                                                    okBtn = questionWin.FindFirst(TreeScope.TreeScope_Descendants, condOk);

                                                    if (okBtn != null)
                                                    {
                                                        bool isEnabled = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                        bool isOffscreen = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                                        if (isEnabled && !isOffscreen)
                                                            break;
                                                        okBtn = null;
                                                    }
                                                }

                                                if (okBtn != null) break;

                                                Thread.Sleep(pollMs);
                                                waited += pollMs;
                                            }

                                            if (okBtn == null)
                                                throw new Exception("Кнопка 'Нет' не найдена или недоступна в окне 'Вопрос'.");

                                            try { okBtn.SetFocus(); } catch { }

                                            bool clicked = false;

                                            try
                                            {
                                                // 1) Пытаемся физическим кликом (твоя функция)
                                                ClickElementWithMouse(okBtn);
                                                clicked = true;
                                            }
                                            catch
                                            {
                                                // 2) InvokePattern
                                                try
                                                {
                                                    var invObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
                                                    if (invObj is IUIAutomationInvokePattern invoke)
                                                    {
                                                        invoke.Invoke();
                                                        clicked = true;
                                                    }
                                                }
                                                catch { }
                                            }

                                            if (!clicked)
                                            {
                                                // 3) LegacyIAccessible
                                                try
                                                {
                                                    var legacyObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId);
                                                    if (legacyObj is IUIAutomationLegacyIAccessiblePattern legacy)
                                                    {
                                                        legacy.DoDefaultAction();
                                                        clicked = true;
                                                    }
                                                }
                                                catch { }
                                            }

                                            if (!clicked)
                                            {
                                                // 4) Fallback: Enter
                                                try
                                                {
                                                    questionWin.SetFocus();
                                                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                                                    clicked = true;
                                                }
                                                catch { }
                                            }

                                            if (!clicked)
                                                throw new Exception("Не удалось нажать кнопку 'Нет'.");

                                            Log(LogLevel.Info, "Кнопка 'Нет' в окне 'Вопрос' успешно нажата.");

                                            // --- ждём закрытие окна "Вопрос" ---
                                            waited = 0;
                                            while (waited < questionTimeoutMs)
                                            {
                                                // Пере-ищем диалог заново, чтобы не полагаться на устаревший COM-указатель
                                                var stillThere = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                                if (stillThere == null)
                                                {
                                                    Log(LogLevel.Info, "Окно 'Вопрос' закрыто (не находится среди потомков DocCard).");
                                                    break;
                                                }

                                                // На некоторых формах диалог может становиться offscreen перед удалением
                                                bool stillEnabled = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                bool stillOffscreen = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;

                                                if (!stillEnabled || stillOffscreen)
                                                {
                                                    Log(LogLevel.Info, "Окно 'Вопрос' больше не активно (disabled/offscreen).");
                                                    break;
                                                }

                                                Thread.Sleep(pollMs);
                                                waited += pollMs;
                                            }

                                            if (waited >= questionTimeoutMs)
                                                throw new Exception("Окно 'Вопрос' не закрылось в отведённое время после нажатия 'ОК'.");
                                        }
                                        catch (Exception ex)
                                        {
                                            Log(LogLevel.Error, $"Ожидание окна внутри MainWindows(Landocs) завершилось ошибкой: {ex.Message}");
                                            throw;
                                        }
                                        #endregion
                                        WaitWindowGoneByHandle(targetWindowAgreement, 3000);
                                        throw new Exception("Договор не проставлен, видимо его нет. Проверьте корректность. Робот завершает работу.");
                                    }
                                }
                                else
                                {
                                    // Если элемент не найден, выбрасываем исключение с детализированным сообщением
                                    SafeCloseWindow(targetWindowAgreement, "Выбор документа");
                                    
                                    #region Выполняем повтоно поиск окна ВОПРОС, тк оно обновилось
                                    try
                                    {
                                        var automation = new CUIAutomation();
                                        var root = automation.GetRootElement();

                                        // --- параметры ожидания ---
                                        const int timeoutMs = 600000; // 10 минут
                                        const int pollMs = 500;       // 0.5 секунды
                                        int waited = 0;

                                        // --- ждём главное окно Landocs ---
                                        IUIAutomationElement mainWin = null;
                                        var condMainAutomationId = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                        var condMainTypeWindow = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                        var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                        while (waited < timeoutMs)
                                        {
                                            var candidates = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                            if (candidates != null && candidates.Length > 0)
                                            {
                                                for (int i = 0; i < candidates.Length; i++)
                                                {
                                                    var e = candidates.GetElement(i);
                                                    var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                    if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        mainWin = e;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (mainWin != null) break;

                                            Thread.Sleep(pollMs);
                                            waited += pollMs;
                                        }

                                        if (mainWin == null)
                                            throw new Exception("Не найдено окно MainWindows (Name содержит 'Landocs') за 10 минут.");

                                        Log(LogLevel.Info, "Главное окно Landocs обнаружено.");

                                        // --- ждём DocCard внутри MainWindows ---
                                        IUIAutomationElement docCard = null;
                                        waited = 0;
                                        var condDocCardId = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                        var condDocTypeWindow = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                        var condDocTypePane = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                        var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                        var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                        while (waited < timeoutMs)
                                        {
                                            docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                            if (docCard != null)
                                            {
                                                var docName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                if (!string.IsNullOrEmpty(docName) && docName.StartsWith("Без имени", StringComparison.Ordinal))
                                                {
                                                    Log(LogLevel.Info, $"DocCard найден: \"{docName}\".");
                                                    break;
                                                }
                                            }
                                            Thread.Sleep(pollMs);
                                            waited += pollMs;
                                        }

                                        if (docCard == null)
                                            throw new Exception("DocCard не найден внутри окна Landocs за 10 минут.");

                                        var finalName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                        if (!finalName.StartsWith("Без имени", StringComparison.Ordinal))
                                            throw new Exception($"Найден DocCard, но Name не начинается с \"Без имени\". Текущее Name: \"{finalName}\".");

                                        // --- ждём окно "Вопрос" внутри DocCard ---
                                        const int questionTimeoutMs = 120000; // 2 мин
                                        waited = 0;
                                        IUIAutomationElement questionWin = null;

                                        var condQuestionName = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_NamePropertyId, "Вопрос");
                                        var condTypeWindow = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                        var condTypePane = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                        var condType = automation.CreateOrCondition(condTypeWindow, condTypePane);
                                        var condQuestion = automation.CreateAndCondition(condQuestionName, condType);

                                        while (waited < questionTimeoutMs)
                                        {
                                            questionWin = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                            if (questionWin != null)
                                            {
                                                bool isEnabled = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                bool isOffscreen = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                                if (isEnabled && !isOffscreen)
                                                    break;
                                                questionWin = null;
                                            }

                                            Thread.Sleep(pollMs);
                                            waited += pollMs;
                                        }

                                        if (questionWin == null)
                                            throw new Exception("Внутри DocCard не найдено доступное окно с Name='Вопрос'.");

                                        try { questionWin.SetFocus(); } catch { }
                                        Log(LogLevel.Info, "Окно 'Вопрос' внутри DocCard найдено и активировано.");

                                        // --- ищем кнопку ОК внутри окна "Вопрос" ---
                                        IUIAutomationElement okBtn = null;
                                        waited = 0;
                                        string[] okNames = { "&НЕТ", "&Нет", "Нет", "нет" }; // оставил твои варианты и добавил частые

                                        var condOkType = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ButtonControlTypeId);

                                        while (waited < questionTimeoutMs && okBtn == null)
                                        {
                                            foreach (var nm in okNames)
                                            {
                                                var condOkName = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, nm);
                                                var condOk = automation.CreateAndCondition(condOkName, condOkType);
                                                okBtn = questionWin.FindFirst(TreeScope.TreeScope_Descendants, condOk);

                                                if (okBtn != null)
                                                {
                                                    bool isEnabled = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                    bool isOffscreen = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                                    if (isEnabled && !isOffscreen)
                                                        break;
                                                    okBtn = null;
                                                }
                                            }

                                            if (okBtn != null) break;

                                            Thread.Sleep(pollMs);
                                            waited += pollMs;
                                        }

                                        if (okBtn == null)
                                            throw new Exception("Кнопка 'Нет' не найдена или недоступна в окне 'Вопрос'.");

                                        try { okBtn.SetFocus(); } catch { }

                                        bool clicked = false;

                                        try
                                        {
                                            // 1) Пытаемся физическим кликом (твоя функция)
                                            ClickElementWithMouse(okBtn);
                                            clicked = true;
                                        }
                                        catch
                                        {
                                            // 2) InvokePattern
                                            try
                                            {
                                                var invObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
                                                if (invObj is IUIAutomationInvokePattern invoke)
                                                {
                                                    invoke.Invoke();
                                                    clicked = true;
                                                }
                                            }
                                            catch { }
                                        }

                                        if (!clicked)
                                        {
                                            // 3) LegacyIAccessible
                                            try
                                            {
                                                var legacyObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId);
                                                if (legacyObj is IUIAutomationLegacyIAccessiblePattern legacy)
                                                {
                                                    legacy.DoDefaultAction();
                                                    clicked = true;
                                                }
                                            }
                                            catch { }
                                        }

                                        if (!clicked)
                                        {
                                            // 4) Fallback: Enter
                                            try
                                            {
                                                questionWin.SetFocus();
                                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                                                clicked = true;
                                            }
                                            catch { }
                                        }

                                        if (!clicked)
                                            throw new Exception("Не удалось нажать кнопку 'Нет'.");

                                        Log(LogLevel.Info, "Кнопка 'Нет' в окне 'Вопрос' успешно нажата.");

                                        // --- ждём закрытие окна "Вопрос" ---
                                        waited = 0;
                                        while (waited < questionTimeoutMs)
                                        {
                                            // Пере-ищем диалог заново, чтобы не полагаться на устаревший COM-указатель
                                            var stillThere = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                            if (stillThere == null)
                                            {
                                                Log(LogLevel.Info, "Окно 'Вопрос' закрыто (не находится среди потомков DocCard).");
                                                break;
                                            }

                                            // На некоторых формах диалог может становиться offscreen перед удалением
                                            bool stillEnabled = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                            bool stillOffscreen = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;

                                            if (!stillEnabled || stillOffscreen)
                                            {
                                                Log(LogLevel.Info, "Окно 'Вопрос' больше не активно (disabled/offscreen).");
                                                break;
                                            }

                                            Thread.Sleep(pollMs);
                                            waited += pollMs;
                                        }

                                        if (waited >= questionTimeoutMs)
                                            throw new Exception("Окно 'Вопрос' не закрылось в отведённое время после нажатия 'ОК'.");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(LogLevel.Error, $"Ожидание окна внутри MainWindows(Landocs) завершилось ошибкой: {ex.Message}");
                                        throw;
                                    }
                                    #endregion
                                    WaitWindowGoneByHandle(targetWindowAgreement, 3000);
                                    throw new Exception("Кнопка [Выбрать] в окне [Выбор документа] со списком журналов не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и ввод подписанта в элемент "Подписант"
                            string xpathSignerInput = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[1]/Pane[13]/Edit";
                            var targetElementSignerInput = FindElementByXPath(targetWindowCreateDoc, xpathSignerInput, 60);

                            if (targetElementSignerInput != null)
                            {
                                string signer = GetConfigValue("Signatory").Trim(); // Получаем значение подписанта из конфигурации
                                string currentSignerInput = targetElementSignerInput.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                if (!string.IsNullOrEmpty(currentSignerInput))
                                {
                                    Log(LogLevel.Info, $"Текущий подписант: [{currentSignerInput}]. Меняю на: [{signer}].");
                                }
                                else
                                {
                                    Log(LogLevel.Info, $"Текущий подписант отсутствует. Устанавливаю нового: [{signer}].");
                                }

                                try
                                {
                                    // Используем ValuePattern для установки значения
                                    if (targetElementSignerInput.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        valuePattern.SetValue(signer);
                                        Log(LogLevel.Info, $"Подписант успешно установлен: [{signer}].");
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент ввода подписанта не поддерживает ValuePattern.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"Ошибка при установке подписанта: {ex.Message}", ex);
                                }
                            }
                            else
                            {
                                throw new Exception("Элемент ввода подписанта не найден. Робот завершает работу.");
                            }
                            #endregion

                            #region Поиск и нажатие кнопки "Сохранить документ"
                            try
                            {
                                string xpathAgreementOkButton = "Pane[2]/Pane/Pane/ToolBar[1]/Button[1]";
                                var targetElementAgreementOkButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementOkButton, 60);

                                if (targetElementAgreementOkButton != null)
                                {
                                    // Устанавливаем фокус на кнопку и нажимаем
                                    targetElementAgreementOkButton.SetFocus();
                                    ClickElementWithMouse(targetElementAgreementOkButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [Сохранить документ] в окне [Создать документ].");
                                }
                                else
                                {
                                    // Выбрасываем исключение, если элемент не найден
                                    throw new Exception("Кнопка [Сохранить документ] в окне [Создать документ] не найдена.");
                                }

                                #region Проверка появления окна ошибки отсутствия контрагента
                                try
                                {
                                    var automation = new CUIAutomation();
                                    var rootElement = automation.GetRootElement();
                                    var targetErrorWindow = FindElementByName(rootElement, "Ошибка", 5);

                                    if (targetErrorWindow != null)
                                    {
                                        var messageCondition = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId,
                                            UIA_ControlTypeIds.UIA_TextControlTypeId);
                                        var messageElement = targetErrorWindow.FindFirst(TreeScope.TreeScope_Descendants, messageCondition);
                                        string messageText = messageElement?.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string ?? string.Empty;

                                        if (!string.IsNullOrEmpty(messageText) && messageText.Contains("Для указанного контрагента отсутствует Соглашение об ЭДО"))
                                        {
                                            Log(LogLevel.Error, "Для указанного контрагента отсутствует Соглашение об ЭДО. Отправка исходящего электронного документа для данного контрагента невозможна. Закрываю приложение.");

                                            var okButton = FindElementByName(targetErrorWindow, "&ОК", 5);
                                            if (okButton != null)
                                            {
                                                TryInvokeElement(okButton);
                                                Thread.Sleep(1000);
                                            }

                                            if (!string.IsNullOrEmpty(landocsProcessName))
                                            {
                                                KillExcelProcesses(landocsProcessName);
                                            }

                                            throw new Exception("Для указанного контрагента отсутствует Соглашение об ЭДО. Отправка исходящего электронного документа для данного контрагента невозможна. Работа робота завершена.");
                                        }

                                        if (!string.IsNullOrEmpty(messageText) && messageText.Contains("Для сохранения данных документа необходимо заполнить поля:"))
                                        {
                                            Log(LogLevel.Error, $"Не заполнено обязательное поле. Текст ошибки:{messageText}");

                                            var okButton = FindElementByName(targetErrorWindow, "&ОК", 5);
                                            if (okButton != null)
                                            {
                                                TryInvokeElement(okButton);
                                                Thread.Sleep(1000);
                                            }

                                            if (!string.IsNullOrEmpty(landocsProcessName))
                                            {
                                                KillExcelProcesses(landocsProcessName);
                                            }

                                            throw new Exception($"Не заполнено обязательное поле. Текст ошибки:{messageText}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                    throw;
                                }
                                #endregion

                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Выполняем повтоно поиск окна документа, тк оно обновилось
                            try
                            {
                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                // --- параметры ожидания ---
                                const int timeoutMs = 600000; // 10 минут
                                const int pollMs = 500;       // 0.5 секунды
                                int waited = 0;

                                // --- ждём главное окно Landocs ---
                                IUIAutomationElement mainWin = null;
                                var condMainAutomationId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                var condMainTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                while (waited < timeoutMs)
                                {
                                    var candidates = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                    if (candidates != null && candidates.Length > 0)
                                    {
                                        for (int i = 0; i < candidates.Length; i++)
                                        {
                                            var e = candidates.GetElement(i);
                                            var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                mainWin = e;
                                                break;
                                            }
                                        }
                                    }
                                    if (mainWin != null) break;

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (mainWin == null)
                                    throw new Exception("Не найдено окно MainWindows (Name содержит 'Landocs') за 10 минут.");

                                Log(LogLevel.Info, "Главное окно Landocs обнаружено.");

                                // --- ждём DocCard внутри MainWindows ---
                                IUIAutomationElement docCard = null;
                                waited = 0;
                                var condDocCardId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                var condDocTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condDocTypePane = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                while (waited < timeoutMs)
                                {
                                    docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                    if (docCard != null)
                                    {
                                        var docName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                        if (!string.IsNullOrEmpty(docName) && docName.StartsWith("№ППУД", StringComparison.Ordinal))
                                        {
                                            Log(LogLevel.Info, $"DocCard найден: \"{docName}\".");
                                            break;
                                        }
                                    }
                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (docCard == null)
                                    throw new Exception("DocCard не найден внутри окна Landocs за 10 минут.");

                                var finalName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                if (!finalName.StartsWith("№ППУД", StringComparison.Ordinal))
                                    throw new Exception($"Найден DocCard, но Name не начинается с \"№ППУД\". Текущее Name: \"{finalName}\".");
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ожидание DocCard в MainWindows(Landocs) завершилось ошибкой: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и клик на вкладку "Структура папок"
                            try
                            {
                                string xpathStructurekFolderTab = "Tab/Pane/Pane/Pane/Tab";
                                var targetElementStructurekFolderTab = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderTab, 60);

                                if (targetElementStructurekFolderTab != null)
                                {
                                    // Поиск элемента "Панель данных"
                                    var targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Структура папок", 60);

                                    int retryCount = 0;
                                    bool isEnabled = false;

                                    // Проверка на доступность элемента
                                    while (targetElementStructurekFolderItem != null && retryCount < 3)
                                    {
                                        isEnabled = (bool)targetElementStructurekFolderItem.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId);

                                        if (isEnabled)
                                        {
                                            break;
                                        }

                                        Log(LogLevel.Info, "Элемент неактивен, ждем 1 минуту...");
                                        Thread.Sleep(60000); // Ждем 1 минуту
                                        targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Структура папок", 60); // Переходим к следующей попытке
                                        retryCount++;
                                    }

                                    if (isEnabled)
                                    {
                                        // Получаем паттерн SelectionItemPattern
                                        if (targetElementStructurekFolderItem.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern SelectionItemPattern)
                                        {
                                            SelectionItemPattern.Select();
                                            ClickElementWithMouse(targetElementStructurekFolderItem);
                                            Log(LogLevel.Info, "Элемент [Структура папок] выбран.");
                                        }
                                        else
                                        {
                                            // Если паттерн не доступен, выбрасываем исключение
                                            throw new Exception("Паттерн SelectionItemPattern не поддерживается для элемента [Структура папок].");
                                        }
                                    }
                                    else
                                    {
                                        // Если элемент неактивен после 3 попыток, выбрасываем исключение
                                        throw new Exception("Элемент [Структура папок] не активен после 3 попыток.");
                                    }

                                    // Устанавливаем фокус на кнопку и нажимаем
                                    targetElementStructurekFolderTab.SetFocus();
                                    //TryInvokeElement(targetElementStructurekFolderTab);
                                    Log(LogLevel.Info, "Нажали на кнопку [Структура папок] в окне [Создать документ].");
                                }
                                else
                                {
                                    // Выбрасываем исключение, если элемент не найден
                                    throw new Exception("Кнопка [Структура папок] в окне [Создать документ] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и проверка дерева "Структуры папок"
                            try
                            {
                                string xpathStructurekFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                var targetElementStructurekFolderTList = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderList, 60);

                                if (targetElementStructurekFolderTList != null)
                                {
                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructurekFolderTList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox != null)
                                    {
                                        // Проверяем, является ли элемент CheckBox
                                        var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                        if (togglePattern != null)
                                        {
                                            // Устанавливаем значение CheckBox на true, если оно не выбрано
                                            if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                            {
                                                togglePattern.Toggle();
                                                Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                            }
                                            else
                                            {
                                                Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                            }

                                            // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                            Thread.Sleep(1000);

                                            // Ищем элемент "Акты сверки" после раскрытия
                                            var checkBoxElementItem = FindElementByName(targetElementStructurekFolderTList, "Акт сверки", 60);

                                            if (checkBoxElementItem != null)
                                            {
                                                // Устанавливаем фокус на элемент и активируем его
                                                checkBoxElementItem.SetFocus();
                                                //TryInvokeElement(checkBoxElementItem);
                                                Log(LogLevel.Info, "Выбран элемент 'Акты сверки' после раскрытия CheckBox.");
                                            }
                                            else
                                            {
                                                throw new Exception("Элемент 'Акты сверки' не найден.");
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Не удалось найти первый дочерний элемент.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Структура папок] не найден в окне [Создать документ].");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и нажатие кнопки "Добавить"
                            try
                            {
                                const string dialogName = "Выберете файлы для прикрепления к РК";
                                const int maxAttempts = 3;
                                const int waitCloseMs = 15000;
                                const int pollMs = 500;

                                bool dialogClosed = false;

                                for (int attempt = 1; attempt <= maxAttempts && !dialogClosed; attempt++)
                                {
                                    Log(LogLevel.Info, $"Старт попытки прикрепления файла (попытка {attempt} из {maxAttempts}).");

                                    // --- "Добавить" ---
                                    var addPanel = FindElementByXPath(targetWindowCreateDoc, "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Pane[6]", 60);
                                    if (addPanel == null) throw new Exception("Панель для кнопки [Добавить] не найдена.");
                                    var addBtn = FindElementByName(addPanel, "Добавить", 60);
                                    if (addBtn == null) throw new Exception("Кнопка [Добавить] не найдена.");
                                    try { addBtn.SetFocus(); } catch { }
                                    ClickElementWithMouse(addBtn);
                                    Log(LogLevel.Info, "Нажали [Добавить].");

                                    // --- Окно выбора ---
                                    var fileDlg = FindElementByName(targetWindowCreateDoc, dialogName, 60);
                                    if (fileDlg == null) throw new Exception($"Окно [{dialogName}] не найдено.");
                                    Log(LogLevel.Info, $"Появилось окно: [{dialogName}].");

                                    // --- Ввод пути ---
                                    var fileNameEdit = FindElementByXPath(fileDlg, "ComboBox[1]/Edit", 60);
                                    if (fileNameEdit == null) throw new Exception("Поле ввода пути не найдено.");
                                    var vp = fileNameEdit.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
                                    if (vp == null) throw new Exception("Поле ввода пути не поддерживает ValuePattern.");
                                    vp.SetValue(filePdf.Trim());
                                    Log(LogLevel.Info, "Путь к файлу установлен.");

                                    // --- Фокус на Open/Открыть и Enter ---
                                    var openBtn = FindElementByName(fileDlg, "Open", 5) ?? FindElementByName(fileDlg, "Открыть", 5);
                                    if (openBtn == null) throw new Exception("Кнопка [Open/Открыть] не найдена.");
                                    try { openBtn.SetFocus(); } catch { }
                                    TryBringToFront(fileDlg); // если используешь — из твоих хелперов
                                    PressEnter();             // если используешь — из твоих хелперов

                                    Log(LogLevel.Info, "Нажал Enter на [Open/Открыть]. Ожидаю закрытие окна…");

                                    // --- Ждём закрытия окна ---
                                    if (WaitWindowClosedByName(targetWindowCreateDoc, dialogName, waitCloseMs, pollMs))
                                    {
                                        Log(LogLevel.Info, "Окно выбора файлов закрылось. Продолжаю дальше.");
                                        dialogClosed = true;           // <--- ВАЖНО: флаг вместо break
                                                                       // НЕ НУЖНО break; цикл сам завершится по условию (&& !dialogClosed)
                                    }
                                    else
                                    {
                                        // Закрываем принудительно и переходим к следующей попытке
                                        Log(LogLevel.Warning, $"Окно [{dialogName}] не закрылось. Пробую закрыть принудительно и повторить…");
                                        ForceCloseWindowByName(targetWindowCreateDoc, dialogName); // из твоих хелперов
                                        Thread.Sleep(700);
                                        // здесь просто пойдём на следующую итерацию for
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при прикреплении pdf: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и проверка дерева "Структуры папок" и проверка что файл был прикреплен
                            try
                            {
                                string xpathStructurekFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                var targetElementStructurekFolderTList = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderList, 60);

                                if (targetElementStructurekFolderTList != null)
                                {
                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructurekFolderTList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox != null)
                                    {
                                        // Проверяем, является ли элемент CheckBox
                                        var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                        if (togglePattern != null)
                                        {
                                            // Устанавливаем значение CheckBox на true, если оно не выбрано
                                            if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                            {
                                                togglePattern.Toggle();
                                                Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                            }
                                            else
                                            {
                                                Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                            }

                                            // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                            Thread.Sleep(1000);

                                            // Ищем элемент "Акты сверки" после раскрытия
                                            var checkBoxElementItem = FindElementByName(targetElementStructurekFolderTList, "Акт сверки", 60);

                                            if (checkBoxElementItem != null)
                                            {
                                                //TODO: Сделать проверку что файл был прикреплен
                                                // Устанавливаем фокус на элемент и активируем его
                                                checkBoxElementItem.SetFocus();
                                                //TryInvokeElement(checkBoxElementItem);
                                                Log(LogLevel.Info, "Выбран элемент 'Акты сверки' после раскрытия CheckBox.");


                                            }
                                            else
                                            {
                                                throw new Exception("Элемент 'Акты сверки' не найден.");
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Не удалось найти первый дочерний элемент.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Структура папок] не найден в окне [Создать документ].");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и проверка дерева "Структуры папок" и проверка что файл был прикреплен
                            try
                            {
                                string xpathStructureFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                int maxRetries = 3;
                                int waitTime = 10000; // 10 секунд между проверками
                                int outerRetryCount = 0;

                                IUIAutomationElement targetElementStructureFolderList = null;

                                while (outerRetryCount < maxRetries)
                                {
                                    // Поиск элемента [Структура папок]
                                    targetElementStructureFolderList = FindElementByXPath(targetWindowCreateDoc, xpathStructureFolderList, 60);

                                    if (targetElementStructureFolderList == null)
                                    {
                                        outerRetryCount++;
                                        Log(LogLevel.Info, $"Элемент [Структура папок] не найден. Попытка {outerRetryCount}/{maxRetries}. Ожидаем обновления интерфейса...");
                                        Thread.Sleep(waitTime);
                                        continue;
                                    }

                                    Log(LogLevel.Info, "Элемент [Структура папок] найден.");

                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructureFolderList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox == null)
                                        throw new Exception("Не удалось найти первый дочерний элемент.");

                                    // Проверяем, является ли элемент CheckBox
                                    var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                    if (togglePattern == null)
                                        throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");

                                    // Устанавливаем значение CheckBox на true, если оно не выбрано
                                    if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                    {
                                        togglePattern.Toggle();
                                        Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                    }

                                    // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                    Thread.Sleep(2000);

                                    // Ищем элемент "Акты сверки" после раскрытия
                                    var checkBoxElementItem = FindElementByName(targetElementStructureFolderList, "Акт сверки", 60);

                                    if (checkBoxElementItem == null)
                                        throw new Exception("Элемент 'Акты сверки' не найден.");

                                    int innerRetryCount = 0;
                                    bool isEnabled = false;
                                    var childrenFilePdf = checkBoxElementItem.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());
                                    bool selectedFile = false;

                                    // Цикл для проверки наличия и доступности прикрепленного файла
                                    while (innerRetryCount < maxRetries)
                                    {
                                        if (childrenFilePdf != null && (bool)childrenFilePdf.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId))
                                        {
                                            isEnabled = true;
                                            break;
                                        }

                                        innerRetryCount++;
                                        Log(LogLevel.Info, $"Прикрепленный файл не найден или неактивен. Попытка {innerRetryCount}/{maxRetries}. Ожидаем 10 секунд...");
                                        Thread.Sleep(waitTime);

                                        // Повторно ищем дочерний элемент
                                        childrenFilePdf = checkBoxElementItem.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());
                                    }

                                    // Если файл найден, завершаем внешний цикл
                                    if (isEnabled)
                                    {
                                        // Получаем значение имени файла
                                        var elementValue = childrenFilePdf.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString() ?? string.Empty;

                                        if (!string.IsNullOrEmpty(elementValue))
                                        {
                                            ClickElementWithMouse(childrenFilePdf);
                                            Log(LogLevel.Info, $"Прикрепленный файл [{elementValue}] выбран.");
                                            selectedFile = true;
                                        }
                                        else
                                        {
                                            throw new Exception("Имя прикрепленного файла не определено.");
                                        }
                                    }

                                    if (selectedFile)
                                    {
                                        // Если файл найден и выбран, выходим из внешнего цикла
                                        break;
                                    }
                                    else
                                    {
                                        // Если файл не найден после внутреннего цикла, переходим к следующей внешней итерации
                                        Log(LogLevel.Info, "Прикрепленный файл не найден после 3 попыток. Повторная проверка элемента [Структура папок].");
                                        outerRetryCount++;
                                    }
                                }

                                // Если внешний цикл завершился безуспешно
                                if (outerRetryCount >= maxRetries)
                                {
                                    throw new Exception("Прикрепленный файл не найден после нескольких проверок.");
                                }

                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [Просмотреть реквизиты документа] для открытия окна редактирования реквизитов документа

                            try
                            {
                                string xpathButtonFileSetting = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Pane[1]/Button";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Просмотреть реквизиты документа] в окне [Создание документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementButtonFileSetting = FindElementByXPath(targetWindowCreateDoc, xpathButtonFileSetting, 10);

                                if (targetElementButtonFileSetting == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Просмотреть реквизиты документа] не найдена в окне [Создание документа].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Просмотреть реквизиты документа] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementButtonFileSetting.SetFocus();
                                    ClickElementWithMouse(targetElementButtonFileSetting);
                                    Log(LogLevel.Info, "Нажали на кнопку [Просмотреть реквизиты документа] в окне [Создание документа].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Просмотреть реквизиты документа].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Просмотреть реквизиты документа]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск окна [Реквизиты документа в архиве]
                            try
                            {
                                // Поиск окна "Выбор документа"
                                targetWindowSettingFile = FindElementByName(targetWindowCreateDoc, "Реквизиты документа в архиве", 60);

                                // Проверка, был ли найден элемент
                                if (targetWindowSettingFile != null)
                                {
                                    Log(LogLevel.Info, "Окно редактирования реквизитов документа [Реквизиты документа в архиве] найдено.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Окно редактирования реквизитов документа [Реквизиты документа в архиве] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске окна
                                Log(LogLevel.Error, $"Ошибка при поиске окна [Реквизиты документа в архиве]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и ввод подписанта в элемент "Подписант"

                            try
                            {
                                string xpathSettingFileName = "Pane/Pane/Tab/Pane/Pane[4]/Pane[7]/Edit";

                                Log(LogLevel.Info, "Начинаем поиск элемента ввода [Номер файла]...");

                                // Поиск элемента ввода
                                var targetSettingFileName = FindElementByXPath(targetWindowSettingFile, xpathSettingFileName, 60);

                                if (targetSettingFileName == null)
                                {
                                    string errorMessage = "Элемент ввода [Номер файла] не найден. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                Log(LogLevel.Info, "Элемент ввода [Номер файла] найден. Получаем текущие значения...");

                                // Получаем значение подписанта из конфигурации
                                string fileName = GetTicketValue("FileNameNumber").Trim();
                                string currentSettingFileName = targetSettingFileName.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                // Логируем текущее значение
                                if (!string.IsNullOrEmpty(currentSettingFileName))
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Номер файла]: [{currentSettingFileName}]. Меняю на: [{fileName}].");
                                }
                                else
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Номер файла] отсутствует. Устанавливаю новое: [{fileName}].");
                                }

                                // Используем ValuePattern для установки значения
                                try
                                {
                                    if (targetSettingFileName.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        valuePattern.SetValue(fileName);
                                        Log(LogLevel.Info, $"Значение [Номер файла] успешно установлено: [{fileName}].");
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент ввода [Номер файла] не поддерживает ValuePattern.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string valueError = $"Ошибка при установке значения [Номер файла]: {ex.Message}";
                                    Log(LogLevel.Error, valueError);
                                    throw new Exception(valueError, ex);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем общую ошибку
                                Log(LogLevel.Error, $"Ошибка при работе с элементом ввода [Номер файла]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и ввод даты документа в элемент [Дата документа]

                            try
                            {
                                string xpathDateInput = "Pane/Pane/Tab/Pane/Pane[4]/Pane[3]/Pane";

                                Log(LogLevel.Info, "Начинаем поиск элемента ввода [Дата документа]...");

                                // Поиск элемента ввода
                                var targetDateInput = FindElementByXPath(targetWindowSettingFile, xpathDateInput, 60);

                                if (targetDateInput == null)
                                {
                                    string errorMessage = "Элемент ввода [Дата документа] не найден. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                Log(LogLevel.Info, "Элемент ввода [Дата документа] найден.");

                                // Устанавливаем фокус на элемент
                                targetDateInput.SetFocus();
                                ClickElementWithMouse(targetDateInput);
                                Log(LogLevel.Info, "Фокус установлен на элемент ввода [Дата документа].");

                                // Получаем текущую дату
                                string day = DateTime.Now.ToString("dd");  // День
                                string month = DateTime.Now.ToString("MM"); // Месяц
                                string year = DateTime.Now.ToString("yyyy"); // Год

                                Log(LogLevel.Info, $"Пытаемся ввести текущую дату: [{day}.{month}.{year}]...");

                                bool dateEntered = false; // Флаг для проверки, введена ли дата
                                try
                                {
                                    SendKeys.SendWait(day);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    SendKeys.SendWait(month);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    SendKeys.SendWait(year);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    Log(LogLevel.Info, $"Дата документа введена в поле ввода.");

                                }
                                catch (Exception ex)
                                {
                                    string inputError = $"Ошибка при вводе даты документа: {ex.Message}";
                                    Log(LogLevel.Error, inputError);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем общую ошибку
                                Log(LogLevel.Error, $"Ошибка при работе с элементом ввода [Дата документа]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и ввод элемента ввода ко-ва эле копий

                            try
                            {
                                string xpathSettingFileCount = "Pane/Pane/Tab/Pane/Pane[4]/Pane[1]/Edit";

                                Log(LogLevel.Info, "Начинаем поиск элемента ввода [Количество копий]...");

                                // Поиск элемента ввода
                                var targetSettingFileCount = FindElementByXPath(targetWindowSettingFile, xpathSettingFileCount, 60);

                                if (targetSettingFileCount == null)
                                {
                                    string errorMessage = "Элемент ввода [Количество копий] не найден. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                Log(LogLevel.Info, "Элемент ввода [Количество копий] найден. Получаем текущие значения...");

                                // Получаем значение подписанта из конфигурации
                                string fileCount = "1";
                                string currentSettingFileName = targetSettingFileCount.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                // Логируем текущее значение
                                if (!string.IsNullOrEmpty(currentSettingFileName))
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Количество копий]: [{currentSettingFileName}]. Меняю на: [{fileCount}].");
                                }
                                else
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Количество копий] отсутствует. Устанавливаю новое: [{fileCount}].");
                                }

                                // Используем ValuePattern для установки значения
                                try
                                {
                                    if (targetSettingFileCount.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        valuePattern.SetValue(fileCount);
                                        Log(LogLevel.Info, $"Значение [Номер файла] успешно установлено: [{fileCount}].");
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент ввода [Номер файла] не поддерживает ValuePattern.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string valueError = $"Ошибка при установке значения [Номер файла]: {ex.Message}";
                                    Log(LogLevel.Error, valueError);
                                    throw new Exception(valueError, ex);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем общую ошибку
                                Log(LogLevel.Error, $"Ошибка при работе с элементом ввода [Номер файла]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Кнопка [Сохранить] для сохранения редактирования реквизитов документа

                            try
                            {
                                string xpathAgreementSaveButton = "Pane/Pane/Tab/Pane/Pane[1]/Button[1]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Сохранить] в окне [Редактирования реквизитов документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementAgreementButton = FindElementByXPath(targetWindowSettingFile, xpathAgreementSaveButton, 10);

                                if (targetElementAgreementButton == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Сохранить] не найдена в окне [Редактирования реквизитов документа].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Сохранить] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementAgreementButton.SetFocus();
                                    ClickElementWithMouse(targetElementAgreementButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [Сохранить] в окне [Редактирования реквизитов документа].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Сохранить].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Сохранить]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и проверка дерева "Структуры папок" и проверка что файл был прикреплен
                            try
                            {
                                string xpathStructureFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                int maxRetries = 3;
                                int waitTime = 10000; // 10 секунд между проверками
                                int outerRetryCount = 0;

                                IUIAutomationElement targetElementStructureFolderList = null;

                                while (outerRetryCount < maxRetries)
                                {
                                    // Поиск элемента [Структура папок]
                                    targetElementStructureFolderList = FindElementByXPath(targetWindowCreateDoc, xpathStructureFolderList, 60);

                                    if (targetElementStructureFolderList == null)
                                    {
                                        outerRetryCount++;
                                        Log(LogLevel.Info, $"Элемент [Структура папок] не найден. Попытка {outerRetryCount}/{maxRetries}. Ожидаем обновления интерфейса...");
                                        Thread.Sleep(waitTime);
                                        continue;
                                    }

                                    Log(LogLevel.Info, "Элемент [Структура папок] найден.");

                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructureFolderList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox == null)
                                        throw new Exception("Не удалось найти первый дочерний элемент.");

                                    // Проверяем, является ли элемент CheckBox
                                    var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                    if (togglePattern == null)
                                        throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");

                                    // Устанавливаем значение CheckBox на true, если оно не выбрано
                                    if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                    {
                                        togglePattern.Toggle();
                                        Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                    }

                                    // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                    Thread.Sleep(2000);

                                    // Ищем элемент "Акты сверки" после раскрытия
                                    var checkBoxElementItem = FindElementByName(targetElementStructureFolderList, "Акт сверки", 60);

                                    if (checkBoxElementItem == null)
                                        throw new Exception("Элемент 'Акты сверки' не найден.");

                                    int innerRetryCount = 0;
                                    bool isEnabled = false;
                                    var childrenFilePdf = checkBoxElementItem.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());
                                    bool selectedFile = false;

                                    // Цикл для проверки наличия и доступности прикрепленного файла
                                    while (innerRetryCount < maxRetries)
                                    {
                                        if (childrenFilePdf != null && (bool)childrenFilePdf.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId))
                                        {
                                            isEnabled = true;
                                            break;
                                        }

                                        innerRetryCount++;
                                        Log(LogLevel.Info, $"Прикрепленный файл не найден или неактивен. Попытка {innerRetryCount}/{maxRetries}. Ожидаем 10 секунд...");
                                        Thread.Sleep(waitTime);

                                        // Повторно ищем дочерний элемент
                                        childrenFilePdf = checkBoxElementItem.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());
                                    }

                                    // Если файл найден, завершаем внешний цикл
                                    if (isEnabled)
                                    {
                                        // Получаем значение имени файла
                                        var elementValue = childrenFilePdf.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString() ?? string.Empty;

                                        if (!string.IsNullOrEmpty(elementValue))
                                        {
                                            ClickElementWithMouse(childrenFilePdf);
                                            Log(LogLevel.Info, $"Прикрепленный файл [{elementValue}] выбран.");
                                            selectedFile = true;
                                        }
                                        else
                                        {
                                            throw new Exception("Имя прикрепленного файла не определено.");
                                        }
                                    }

                                    if (selectedFile)
                                    {
                                        // Если файл найден и выбран, выходим из внешнего цикла
                                        break;
                                    }
                                    else
                                    {
                                        // Если файл не найден после внутреннего цикла, переходим к следующей внешней итерации
                                        Log(LogLevel.Info, "Прикрепленный файл не найден после 3 попыток. Повторная проверка элемента [Структура папок].");
                                        outerRetryCount++;
                                    }
                                }

                                // Если внешний цикл завершился безуспешно
                                if (outerRetryCount >= maxRetries)
                                {
                                    throw new Exception("Прикрепленный файл не найден после нескольких проверок.");
                                }

                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [Просмотреть реквизиты документа] для открытия окна редактирования реквизитов документа

                            try
                            {
                                string xpathButtonFileSetting = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Pane[3]/Button[1]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Просмотреть реквизиты документа] в окне [Создание документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementButtonFileSetting = FindElementByXPath(targetWindowCreateDoc, xpathButtonFileSetting, 10);

                                if (targetElementButtonFileSetting == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Просмотреть реквизиты документа] не найдена в окне [Создание документа].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Просмотреть реквизиты документа] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementButtonFileSetting.SetFocus();
                                    ClickElementWithMouse(targetElementButtonFileSetting);
                                    Log(LogLevel.Info, "Нажали на кнопку [Просмотреть реквизиты документа] в окне [Создание документа].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Просмотреть реквизиты документа].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Просмотреть реквизиты документа]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск окна [Реквизиты документа в архиве]
                            try
                            {
                                // Поиск окна "Выбор документа"
                                targetWindowSettingFile2 = FindElementByName(targetWindowCreateDoc, "Реквизиты документа", 60);

                                // Проверка, был ли найден элемент
                                if (targetWindowSettingFile2 != null)
                                {
                                    Log(LogLevel.Info, "Окно редактирования реквизитов документа [Реквизиты документа] найдено.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Окно редактирования реквизитов документа [Реквизиты документа] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске окна
                                Log(LogLevel.Error, $"Ошибка при поиске окна [Реквизиты документа]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [...] для сохранения редактирования реквизитов документа [ЭДО]

                            try
                            {
                                string xpathSettingFileButton = "Pane/Pane/Tab/Pane/Pane/Pane[32]/Edit/Button[1]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [...] в окне [Реквизиты документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementSettingFileButton = FindElementByXPath(targetWindowSettingFile2, xpathSettingFileButton, 10);

                                if (targetElementSettingFileButton == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [...] не найдена в окне [Реквизиты документа].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [...] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementSettingFileButton.SetFocus();
                                    ClickElementWithMouse(targetElementSettingFileButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [...] в окне [Реквизиты документа].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [...] в окне [Реквизиты документа].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [...]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск окна [Выбор элемента] в окне [Реквизиты документа в архиве]
                            try
                            {
                                // Поиск окна "Выбор документа"
                                targetWindowSettingFile2SelectedWinodw = FindElementByName(targetWindowSettingFile2, "Выбор элемента", 60);

                                // Проверка, был ли найден элемент
                                if (targetWindowSettingFile2SelectedWinodw != null)
                                {
                                    Log(LogLevel.Info, "Окно редактирования реквизитов документа [Выбор элемента] найдено.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Окно редактирования реквизитов документа [Выбор элемента] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске окна
                                Log(LogLevel.Error, $"Ошибка при поиске окна [Выбор элемента]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск типа документа в списке
                            try
                            {
                                // Поиск элемента ППУД в списке документов
                                string xpathTypeDocList = "Pane[1]/Pane/Table";
                                Log(LogLevel.Info, "Начинаем поиск элемента 'Список типов документов'...");

                                IUIAutomationElement targetElementDocList = FindElementByXPath(targetWindowSettingFile2SelectedWinodw, xpathTypeDocList, 60);
                                if (targetElementDocList == null)
                                {
                                    throw new Exception("Ошибка: Элемент 'Список типов документов' не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, "Элемент 'Список типов документов' найден. Пытаемся найти 'Панель данных' внутри списка...");
                                IUIAutomationElement dataPanel = FindElementByName(targetElementDocList, "Панель данных", 60);

                                if (dataPanel == null)
                                {
                                    throw new Exception("Ошибка: Элемент 'Панель данных' не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, "'Панель данных' найдена. Получаем cписок типов документов...");
                                IUIAutomationElementArray childrenCounterparty = dataPanel.FindAll(
                                    TreeScope.TreeScope_Children,
                                    new CUIAutomation().CreateTrueCondition()
                                );

                                if (childrenCounterparty == null || childrenCounterparty.Length == 0)
                                {
                                    Log(LogLevel.Warning, "Список типов документов пуст или не найден.");
                                    throw new Exception("Ошибка: Список типов документов пуст или не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, $"Получен cписок типов документов: найдено {childrenCounterparty.Length} элементов.");

                                string typeDocument = "Акт сверки";

                                for (int i = 0; i < childrenCounterparty.Length; i++)
                                {
                                    Log(LogLevel.Debug, $"Обработка типов документов под индексом [{i}]...");

                                    IUIAutomationElement itemCounterparty = childrenCounterparty.GetElement(i);
                                    IUIAutomationElement dataItem = FindElementByXPath(itemCounterparty, "dataitem", 60);

                                    if (dataItem == null)
                                    {
                                        Log(LogLevel.Warning, $"Элемент типов документов под индексом [{i}] не содержит элемента 'dataitem'. Пропускаем...");
                                        continue;
                                    }

                                    if (dataItem.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        string value = valuePattern.CurrentValue ?? string.Empty;
                                        Log(LogLevel.Debug, $"Найден тип документа [{i}]: [{value}]");

                                        // Обрабатываем строки и добавляем в словарь
                                        if (value.ToLower().Trim() == typeDocument.ToLower().Trim())
                                        {
                                            dataItem.SetFocus();
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, $"Элемент типа документа [{i}] не поддерживает ValuePattern. Пропускаем...");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или выборе типа документа в списке результатов: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [Выбрать] для выбора элемента. Тип документа [ЭДО]

                            try
                            {
                                string xpathSettingSelectedWinodwOk = "Pane[2]/Button[1]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Выбрать] в окне [Выбор элементов]...");

                                // Поиск кнопки выбора договора
                                var targetElementSelectedWinodwOk = FindElementByXPath(targetWindowSettingFile2SelectedWinodw, xpathSettingSelectedWinodwOk, 10);

                                if (targetElementSelectedWinodwOk == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Выбрать] не найдена в окне [Выбор элементов].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Выбрать] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementSelectedWinodwOk.SetFocus();
                                    ClickElementWithMouse(targetElementSelectedWinodwOk);
                                    Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор элементов].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Выбрать] в окне [Выбор элементов].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Выбрать]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и ввод элемента номера документа

                            try
                            {
                                string xpathSettingFileCount = "Pane/Pane/Tab/Pane/Pane/Pane[29]/Edit";

                                Log(LogLevel.Info, "Начинаем поиск элемента ввода [Номер документа]...");

                                // Поиск элемента ввода
                                var targetSettingFileCount = FindElementByXPath(targetWindowSettingFile2, xpathSettingFileCount, 60);

                                if (targetSettingFileCount == null)
                                {
                                    string errorMessage = "Элемент ввода [Номер документа] не найден. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                Log(LogLevel.Info, "Элемент ввода [Номер документа] найден. Получаем текущие значения...");

                                // Получаем значение подписанта из конфигурации
                                string fileCount = GetTicketValue("FileNameNumber");
                                string currentSettingFileName = targetSettingFileCount.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                // Логируем текущее значение
                                if (!string.IsNullOrEmpty(currentSettingFileName))
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Номер документа]: [{currentSettingFileName}]. Меняю на: [{fileCount}].");
                                }
                                else
                                {
                                    Log(LogLevel.Info, $"Текущее значение [Номер документа] отсутствует. Устанавливаю новое: [{fileCount}].");
                                }

                                // Используем ValuePattern для установки значения
                                try
                                {
                                    if (targetSettingFileCount.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        valuePattern.SetValue(fileCount);
                                        Log(LogLevel.Info, $"Значение [Номер документа] успешно установлено: [{fileCount}].");
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент ввода [Номер документа] не поддерживает ValuePattern.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string valueError = $"Ошибка при установке значения [Номер документа]: {ex.Message}";
                                    Log(LogLevel.Error, valueError);
                                    throw new Exception(valueError, ex);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем общую ошибку
                                Log(LogLevel.Error, $"Ошибка при работе с элементом ввода [Номер документа]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и ввод даты документа в элемент [Дата документа]

                            try
                            {
                                string xpathDateInput = "Pane/Pane/Tab/Pane/Pane/Pane[27]/Pane";

                                Log(LogLevel.Info, "Начинаем поиск элемента ввода [Дата документа]...");

                                // Поиск элемента ввода
                                var targetDateInput = FindElementByXPath(targetWindowSettingFile2, xpathDateInput, 60);

                                if (targetDateInput == null)
                                {
                                    string errorMessage = "Элемент ввода [Дата документа] не найден. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                Log(LogLevel.Info, "Элемент ввода [Дата документа] найден.");

                                // Устанавливаем фокус на элемент
                                targetDateInput.SetFocus();
                                ClickElementWithMouse(targetDateInput);
                                Log(LogLevel.Info, "Фокус установлен на элемент ввода [Дата документа].");

                                string currentData = GetTicketValue("FileDate");

                                if (string.IsNullOrWhiteSpace(currentData) || !DateTime.TryParse(currentData, out DateTime parsedDate))
                                {
                                    string errorMessage = "Не удалось получить корректную дату из тикета [FileDate]. Робот завершает работу.";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Извлекаем день, месяц и год
                                string day = parsedDate.ToString("dd");
                                string month = parsedDate.ToString("MM");
                                string year = parsedDate.ToString("yyyy");

                                Log(LogLevel.Info, $"Пытаемся ввести текущую дату: [{day}.{month}.{year}]...");
                                try
                                {
                                    SendKeys.SendWait(day);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    SendKeys.SendWait(month);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    SendKeys.SendWait(year);
                                    Thread.Sleep(1000); // Задержка 1 секунда

                                    Log(LogLevel.Info, $"Дата документа введена в поле ввода.");

                                }
                                catch (Exception ex)
                                {
                                    string inputError = $"Ошибка при вводе даты документа: {ex.Message}";
                                    Log(LogLevel.Error, inputError);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем общую ошибку
                                Log(LogLevel.Error, $"Ошибка при работе с элементом ввода [Дата документа]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Кнопка [Выбрать] для реквизиты документа. Окно [Реквизиты документа] [ЭДО]

                            try
                            {
                                string xpathSettingSelectedWinodwOk = "Pane/Pane/Tab/Pane/Pane/Button[1]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Выбрать] в окне [Выбор элементов]...");

                                // Поиск кнопки выбора договора
                                var targetElementSelectedWinodwOk = FindElementByXPath(targetWindowSettingFile2, xpathSettingSelectedWinodwOk, 10);

                                if (targetElementSelectedWinodwOk == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Выбрать] не найдена в окне [Выбор элементов].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Выбрать] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementSelectedWinodwOk.SetFocus();
                                    ClickElementWithMouse(targetElementSelectedWinodwOk);
                                    Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор элементов].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Выбрать] в окне [Выбор элементов].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Выбрать]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и клик на вкладку "Согласование"
                            try
                            {
                                string xpathStructurekFolderTab = "Tab/Pane/Pane/Pane/Tab";
                                var targetElementStructurekFolderTab = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderTab, 60);

                                if (targetElementStructurekFolderTab != null)
                                {
                                    // Поиск элемента "Панель данных"
                                    var targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Согласование", 60);

                                    int retryCount = 0;
                                    bool isEnabled = false;

                                    // Проверка на доступность элемента
                                    while (targetElementStructurekFolderItem != null && retryCount < 3)
                                    {
                                        isEnabled = (bool)targetElementStructurekFolderItem.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId);

                                        if (isEnabled)
                                        {
                                            break;
                                        }

                                        Log(LogLevel.Info, "Элемент неактивен, ждем 1 минуту...");
                                        Thread.Sleep(60000); // Ждем 1 минуту
                                        targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Согласование", 60); // Переходим к следующей попытке
                                        retryCount++;
                                    }

                                    if (isEnabled)
                                    {
                                        // Получаем паттерн SelectionItemPattern
                                        if (targetElementStructurekFolderItem.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern SelectionItemPattern)
                                        {
                                            SelectionItemPattern.Select();
                                            ClickElementWithMouse(targetElementStructurekFolderItem);
                                            Log(LogLevel.Info, "Элемент [Согласование] выбран.");
                                        }
                                        else
                                        {
                                            // Если паттерн не доступен, выбрасываем исключение
                                            throw new Exception("Паттерн SelectionItemPattern не поддерживается для элемента [Согласование].");
                                        }
                                    }
                                    else
                                    {
                                        // Если элемент неактивен после 3 попыток, выбрасываем исключение
                                        throw new Exception("Элемент [Согласование] не активен после 3 попыток.");
                                    }

                                    // Устанавливаем фокус на кнопку и нажимаем
                                    targetElementStructurekFolderTab.SetFocus();
                                    //TryInvokeElement(targetElementStructurekFolderTab);
                                    Log(LogLevel.Info, "Нажали на кнопку [Согласование] в окне [Создать документ].");
                                }
                                else
                                {
                                    // Выбрасываем исключение, если элемент не найден
                                    throw new Exception("Кнопка [Согласование] в окне [Создать документ] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Кнопка [Отправить на маршрут]. Вкладка [Согласование] [ЭДО]

                            try
                            {
                                string xpathSettingSelectedWinodwOk = "Tab/Pane/Pane/Pane/Tab/Pane/Button[11]";

                                Log(LogLevel.Info, "Начинаем поиск кнопки [Отправить на маршрут] в окне [Согласование]...");

                                // Поиск кнопки выбора договора
                                var targetElementSelectedWinodwOk = FindElementByXPath(targetWindowCreateDoc, xpathSettingSelectedWinodwOk, 10);

                                if (targetElementSelectedWinodwOk == null)
                                {
                                    // Если кнопка не найдена, выбрасываем исключение
                                    string errorMessage = "Кнопка [Отправить на маршрут] не найдена в окне [Согласование].";
                                    Log(LogLevel.Error, errorMessage);
                                    throw new Exception(errorMessage);
                                }

                                // Если кнопка найдена, логируем и выполняем действия
                                Log(LogLevel.Info, "Кнопка [Отправить на маршрут] найдена. Пытаемся нажать на кнопку...");

                                // Попытка установить фокус и нажать на кнопку
                                try
                                {
                                    targetElementSelectedWinodwOk.SetFocus();
                                    ClickElementWithMouse(targetElementSelectedWinodwOk);
                                    Log(LogLevel.Info, "Нажали на кнопку [Отправить на маршрут] в окне [Согласование].");
                                }
                                catch (Exception clickEx)
                                {
                                    string clickErrorMessage = "Не удалось нажать на кнопку [Отправить на маршрут] в окне [Согласование].";
                                    Log(LogLevel.Error, $"{clickErrorMessage} Ошибка: {clickEx.Message}");
                                    throw new Exception(clickErrorMessage, clickEx);
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку верхнего уровня и выбрасываем исключение
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Отправить на маршрут]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Выполняем повтоно поиск окна ВОПРОС, тк оно обновилось
                            try
                            {
                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                // --- параметры ожидания ---
                                const int timeoutMs = 600000; // 10 минут
                                const int pollMs = 500;       // 0.5 секунды
                                int waited = 0;

                                // --- ждём главное окно Landocs ---
                                IUIAutomationElement mainWin = null;
                                var condMainAutomationId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                var condMainTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                while (waited < timeoutMs)
                                {
                                    var candidates = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                    if (candidates != null && candidates.Length > 0)
                                    {
                                        for (int i = 0; i < candidates.Length; i++)
                                        {
                                            var e = candidates.GetElement(i);
                                            var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                mainWin = e;
                                                break;
                                            }
                                        }
                                    }
                                    if (mainWin != null) break;

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (mainWin == null)
                                    throw new Exception("Не найдено окно MainWindows (Name содержит 'Landocs') за 10 минут.");

                                Log(LogLevel.Info, "Главное окно Landocs обнаружено.");

                                // --- ждём DocCard внутри MainWindows ---
                                IUIAutomationElement docCard = null;
                                waited = 0;
                                var condDocCardId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                var condDocTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condDocTypePane = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                while (waited < timeoutMs)
                                {
                                    docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                    if (docCard != null)
                                    {
                                        var docName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                        if (!string.IsNullOrEmpty(docName) && docName.StartsWith("№ППУД", StringComparison.Ordinal))
                                        {
                                            Log(LogLevel.Info, $"DocCard найден: \"{docName}\".");
                                            break;
                                        }
                                    }
                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (docCard == null)
                                    throw new Exception("DocCard не найден внутри окна Landocs за 10 минут.");

                                var finalName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                if (!finalName.StartsWith("№ППУД", StringComparison.Ordinal))
                                    throw new Exception($"Найден DocCard, но Name не начинается с \"№ППУД\". Текущее Name: \"{finalName}\".");

                                // --- ждём окно "Вопрос" внутри DocCard ---
                                const int questionTimeoutMs = 120000; // 2 мин
                                waited = 0;
                                IUIAutomationElement questionWin = null;

                                var condQuestionName = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_NamePropertyId, "Вопрос");
                                var condTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condTypePane = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                var condType = automation.CreateOrCondition(condTypeWindow, condTypePane);
                                var condQuestion = automation.CreateAndCondition(condQuestionName, condType);

                                while (waited < questionTimeoutMs)
                                {
                                    questionWin = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                    if (questionWin != null)
                                    {
                                        bool isEnabled = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                        bool isOffscreen = (questionWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                        if (isEnabled && !isOffscreen)
                                            break;
                                        questionWin = null;
                                    }

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (questionWin == null)
                                    throw new Exception("Внутри DocCard не найдено доступное окно с Name='Вопрос'.");

                                try { questionWin.SetFocus(); } catch { }
                                Log(LogLevel.Info, "Окно 'Вопрос' внутри DocCard найдено и активировано.");

                                // --- ищем кнопку ОК внутри окна "Вопрос" ---
                                IUIAutomationElement okBtn = null;
                                waited = 0;
                                string[] okNames = { "&ОК", "ОК", "OK", "Ок", "Ok" }; // оставил твои варианты и добавил частые

                                var condOkType = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ButtonControlTypeId);

                                while (waited < questionTimeoutMs && okBtn == null)
                                {
                                    foreach (var nm in okNames)
                                    {
                                        var condOkName = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, nm);
                                        var condOk = automation.CreateAndCondition(condOkName, condOkType);
                                        okBtn = questionWin.FindFirst(TreeScope.TreeScope_Descendants, condOk);

                                        if (okBtn != null)
                                        {
                                            bool isEnabled = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                            bool isOffscreen = (okBtn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                            if (isEnabled && !isOffscreen)
                                                break;
                                            okBtn = null;
                                        }
                                    }

                                    if (okBtn != null) break;

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (okBtn == null)
                                    throw new Exception("Кнопка 'ОК' не найдена или недоступна в окне 'Вопрос'.");

                                try { okBtn.SetFocus(); } catch { }

                                bool clicked = false;

                                try
                                {
                                    // 1) Пытаемся физическим кликом (твоя функция)
                                    ClickElementWithMouse(okBtn);
                                    clicked = true;
                                }
                                catch
                                {
                                    // 2) InvokePattern
                                    try
                                    {
                                        var invObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
                                        if (invObj is IUIAutomationInvokePattern invoke)
                                        {
                                            invoke.Invoke();
                                            clicked = true;
                                        }
                                    }
                                    catch { }
                                }

                                if (!clicked)
                                {
                                    // 3) LegacyIAccessible
                                    try
                                    {
                                        var legacyObj = okBtn.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId);
                                        if (legacyObj is IUIAutomationLegacyIAccessiblePattern legacy)
                                        {
                                            legacy.DoDefaultAction();
                                            clicked = true;
                                        }
                                    }
                                    catch { }
                                }

                                if (!clicked)
                                {
                                    // 4) Fallback: Enter
                                    try
                                    {
                                        questionWin.SetFocus();
                                        System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                                        clicked = true;
                                    }
                                    catch { }
                                }

                                if (!clicked)
                                    throw new Exception("Не удалось нажать кнопку 'ОК'.");

                                Log(LogLevel.Info, "Кнопка 'ОК' в окне 'Вопрос' успешно нажата.");

                                // --- ждём закрытие окна "Вопрос" ---
                                waited = 0;
                                while (waited < questionTimeoutMs)
                                {
                                    // Пере-ищем диалог заново, чтобы не полагаться на устаревший COM-указатель
                                    var stillThere = docCard.FindFirst(TreeScope.TreeScope_Descendants, condQuestion);
                                    if (stillThere == null)
                                    {
                                        Log(LogLevel.Info, "Окно 'Вопрос' закрыто (не находится среди потомков DocCard).");
                                        break;
                                    }

                                    // На некоторых формах диалог может становиться offscreen перед удалением
                                    bool stillEnabled = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                    bool stillOffscreen = (stillThere.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;

                                    if (!stillEnabled || stillOffscreen)
                                    {
                                        Log(LogLevel.Info, "Окно 'Вопрос' больше не активно (disabled/offscreen).");
                                        break;
                                    }

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (waited >= questionTimeoutMs)
                                    throw new Exception("Окно 'Вопрос' не закрылось в отведённое время после нажатия 'ОК'.");
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ожидание окна внутри MainWindows(Landocs) завершилось ошибкой: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Обновляем MainWindow → ищем DocCard → ждём lblStaticText "Количество запусков: ... Маршрут запущен"
                            try
                            {
                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                const int timeoutMs = 300000; // 5 минут
                                const int pollMs = 5000;      // 5 секунд

                                int waited = 0;

                                // --- условия для MainWindow ---
                                var condMainAutomationId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                var condMainTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                // --- условия для DocCard ---
                                var condDocCardId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                var condDocTypeWindow = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                var condDocTypePane = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                // --- условия для lblStaticText ---
                                var condLblId = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_AutomationIdPropertyId, "lblStaticText");
                                var condLblTypeText = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_TextControlTypeId);
                                var condLblTypeEdit = automation.CreatePropertyCondition(
                                    UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_EditControlTypeId);
                                var condLblType = automation.CreateOrCondition(condLblTypeText, condLblTypeEdit);
                                var condLbl = automation.CreateAndCondition(condLblId, condLblType);

                                const string targetExactName = "Количество запусков: 1. Маршрут запущен:";

                                IUIAutomationElement lbl = null;

                                while (waited < timeoutMs)
                                {
                                    // 1) Находим MainWindow
                                    IUIAutomationElement mainWin = null;
                                    var mains = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                    if (mains != null && mains.Length > 0)
                                    {
                                        for (int i = 0; i < mains.Length; i++)
                                        {
                                            var e = mains.GetElement(i);
                                            var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                mainWin = e;
                                                break;
                                            }
                                        }
                                    }

                                    if (mainWin != null)
                                    {
                                        // 2) Находим DocCard
                                        var docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                        if (docCard != null)
                                        {
                                            var docName = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (!string.IsNullOrEmpty(docName) && docName.StartsWith("№ППУД", StringComparison.Ordinal))
                                            {
                                                // 3) Ищем lblStaticText
                                                lbl = docCard.FindFirst(TreeScope.TreeScope_Descendants, condLbl);
                                                if (lbl != null)
                                                {
                                                    var text = (lbl.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                    bool ok =
                                                        string.Equals(text, targetExactName, StringComparison.Ordinal) ||
                                                        (text.StartsWith("Количество запусков:", StringComparison.Ordinal) &&
                                                         text.IndexOf("Маршрут запущен", StringComparison.OrdinalIgnoreCase) >= 0);

                                                    if (ok)
                                                    {
                                                        Log(LogLevel.Info, $"Найден lblStaticText с требуемым текстом: \"{text}\".");
                                                        break; // успех — выходим из while
                                                    }
                                                    else
                                                    {
                                                        Log(LogLevel.Debug, $"lblStaticText найден, но текст пока другой: \"{text}\".");
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                if (lbl == null || waited >= timeoutMs)
                                    throw new Exception("lblStaticText с текстом 'Количество запусков: ... Маршрут запущен' не появился за 5 минут.");
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске lblStaticText в DocCard: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region MainWindow → DocCard → Lower Ribbon → "Сохранить и закрыть" (клик мышью) + ожидание закрытия DocCard
                            try
                            {
                                // --- параметры ожидания ---
                                const int findTimeoutMs = 120000; // 2 мин: поиск панели и кнопки
                                const int closeTimeoutMs = 300000; // 5 мин: закрытие DocCard
                                const int pollMs = 500;

                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                // --- 1) Находим главное окно LanDocs (MainWindow) ---
                                IUIAutomationElement mainWin = null;
                                {
                                    var condMainAutomationId = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                    var condMainTypeWindow = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                    var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                    int waited = 0;
                                    while (waited < findTimeoutMs)
                                    {
                                        var mains = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                        if (mains != null && mains.Length > 0)
                                        {
                                            for (int i = 0; i < mains.Length; i++)
                                            {
                                                var e = mains.GetElement(i);
                                                var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                                if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    mainWin = e;
                                                    break;
                                                }
                                            }
                                        }
                                        if (mainWin != null) break;

                                        Thread.Sleep(pollMs);
                                        waited += pollMs;
                                    }
                                    if (mainWin == null) throw new Exception("MainWindow (LanDocs) не найден.");
                                }

                                // --- 2) Находим DocCard внутри MainWindow ---
                                IUIAutomationElement docCard = null;
                                {
                                    var condDocCardId = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                    var condDocTypeWindow = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                    var condDocTypePane = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                    var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                    var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                    int waited = 0;
                                    while (waited < findTimeoutMs)
                                    {
                                        docCard = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                        if (docCard != null)
                                        {
                                            bool isEnabled = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                            bool isOffscreen = (docCard.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                            if (isEnabled && !isOffscreen) break;
                                            docCard = null;
                                        }
                                        Thread.Sleep(pollMs);
                                        waited += pollMs;
                                    }
                                    if (docCard == null) throw new Exception("DocCard не найден или недоступен.");
                                }

                                // --- 3) Находим панель "Lower Ribbon" внутри DocCard ---
                                IUIAutomationElement lowerRibbon = null;
                                {
                                    var condLowerName = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_NamePropertyId, "Lower Ribbon");
                                    var condLowerPane = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                    var condLower = automation.CreateAndCondition(condLowerName, condLowerPane);

                                    int waited = 0;
                                    while (waited < findTimeoutMs)
                                    {
                                        lowerRibbon = docCard.FindFirst(TreeScope.TreeScope_Descendants, condLower);
                                        if (lowerRibbon != null)
                                        {
                                            bool isEnabled = (lowerRibbon.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                            bool isOffscreen = (lowerRibbon.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                            if (isEnabled && !isOffscreen) break;
                                            lowerRibbon = null;
                                        }
                                        Thread.Sleep(pollMs);
                                        waited += pollMs;
                                    }
                                    if (lowerRibbon == null) throw new Exception("Панель 'Lower Ribbon' не найдена или недоступна.");
                                }

                                // --- 4) Находим кнопку "Сохранить и закрыть" и жмём физически ---
                                IUIAutomationElement saveCloseBtn = null;
                                {
                                    string[] btnNames =
                                    {
                                        "Сохранить и закрыть",
                                        "&Сохранить и закрыть",
                                        "С&охранить и закрыть",
                                        "Сохранить и закрыть "
                                    };

                                    var condBtnType = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ButtonControlTypeId);

                                    int waited = 0;
                                    while (waited < findTimeoutMs && saveCloseBtn == null)
                                    {
                                        foreach (var nm in btnNames)
                                        {
                                            var condBtnName = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, nm);
                                            var condBtn = automation.CreateAndCondition(condBtnName, condBtnType);

                                            var btn = lowerRibbon.FindFirst(TreeScope.TreeScope_Descendants, condBtn);
                                            if (btn != null)
                                            {
                                                bool isEnabled = (btn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                                bool isOffscreen = (btn.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;
                                                if (isEnabled && !isOffscreen)
                                                {
                                                    saveCloseBtn = btn;
                                                    break;
                                                }
                                            }
                                        }

                                        if (saveCloseBtn != null) break;
                                        Thread.Sleep(pollMs);
                                        waited += pollMs;
                                    }

                                    if (saveCloseBtn == null)
                                        throw new Exception("Кнопка 'Сохранить и закрыть' не найдена или недоступна на панели 'Lower Ribbon'.");

                                    try { saveCloseBtn.SetFocus(); } catch { }

                                    // Требование: физический клик мышью
                                    try
                                    {
                                        ClickElementWithMouse(saveCloseBtn);
                                        Log(LogLevel.Info, "Нажата кнопка 'Сохранить и закрыть' (физический клик).");
                                    }
                                    catch (Exception clickEx)
                                    {
                                        throw new Exception($"Не удалось кликнуть 'Сохранить и закрыть' физически: {clickEx.Message}");
                                    }
                                }

                                // --- 5) Ждём закрытия DocCard ---
                                {
                                    var condDocCardId = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_AutomationIdPropertyId, "DocCard");
                                    var condDocTypeWindow = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                    var condDocTypePane = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_PaneControlTypeId);
                                    var condDocType = automation.CreateOrCondition(condDocTypeWindow, condDocTypePane);
                                    var condDoc = automation.CreateAndCondition(condDocCardId, condDocType);

                                    int waited = 0;
                                    while (waited < closeTimeoutMs)
                                    {
                                        var stillDoc = mainWin.FindFirst(TreeScope.TreeScope_Descendants, condDoc);
                                        if (stillDoc == null)
                                        {
                                            Log(LogLevel.Info, "DocCard закрыт.");
                                            break;
                                        }

                                        bool enabled = (stillDoc.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId) as bool?) ?? false;
                                        bool offscreen = (stillDoc.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsOffscreenPropertyId) as bool?) ?? true;

                                        if (!enabled || offscreen)
                                        {
                                            Log(LogLevel.Info, "DocCard больше не активен (disabled/offscreen) — считаем закрытым.");
                                            break;
                                        }

                                        Thread.Sleep(pollMs);
                                        waited += pollMs;
                                    }

                                    if (waited >= closeTimeoutMs)
                                        throw new Exception("DocCard не закрылся в отведённое время после нажатия 'Сохранить и закрыть'.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при сценарии 'Сохранить и закрыть' в DocCard: {ex.Message}");
                                throw;
                            }
                            #endregion

                            try
                            {
                                if (!Path.GetFileName(filePdf).StartsWith("(+)", StringComparison.Ordinal))
                                {
                                    string directory = Path.GetDirectoryName(filePdf) ?? string.Empty;
                                    string newFileName = $"(+){Path.GetFileName(filePdf)}";
                                    string newFilePath = string.IsNullOrEmpty(directory)
                                        ? newFileName
                                        : Path.Combine(directory, newFileName);

                                    if (File.Exists(newFilePath))
                                    {
                                        Log(LogLevel.Warning, $"Файл [{newFileName}] уже существует. Пропускаю переименование.");
                                    }
                                    else
                                    {
                                        File.Move(filePdf, newFilePath);
                                        Log(LogLevel.Info, $"Файл [{Path.GetFileName(filePdf)}] переименован в [{newFileName}].");
                                    }
                                }
                            }
                            catch (Exception renameEx)
                            {
                                Log(LogLevel.Error, $"Не удалось переименовать файл [{filePdf}]: {renameEx.Message}");
                                Console.WriteLine("end");
                            }

                        }
                        catch (Exception landocsEx)
                        {
                            Log(LogLevel.Error, $"Ошибка в работе LanDocs [{ticket}]: {landocsEx.Message}");
                            MessageBox.Show($"Ошибка в работе LanDocs [{ticket}]: {landocsEx.Message}");

                            #region Мягкое закрытие MainWindow + проверка → жёсткое завершение при необходимости
                            try
                            {
                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                const int waitCloseMs = 60000; // 1 минута
                                const int pollMs = 500;

                                // --- Находим MainWindow (LanDocs) ---
                                IUIAutomationElement mainWin = null;
                                {
                                    var condMainAutomationId = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                    var condMainTypeWindow = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                    var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                    var mains = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                    if (mains != null && mains.Length > 0)
                                    {
                                        for (int i = 0; i < mains.Length; i++)
                                        {
                                            var e = mains.GetElement(i);
                                            var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                mainWin = e;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (mainWin == null)
                                    throw new Exception("MainWindow (LanDocs) не найден для закрытия.");

                                // --- 1) Мягкое закрытие ---
                                bool softCloseAttempted = false;
                                try
                                {
                                    var winPatternObj = mainWin.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId);
                                    if (winPatternObj is IUIAutomationWindowPattern winPattern)
                                    {
                                        winPattern.Close();
                                        softCloseAttempted = true;
                                        Log(LogLevel.Info, "Попробовали мягко закрыть MainWindow через WindowPattern.Close().");
                                    }
                                }
                                catch
                                {
                                    // Если WindowPattern не сработал, пробуем Alt+F4
                                    try
                                    {
                                        mainWin.SetFocus();
                                        System.Windows.Forms.SendKeys.SendWait("%{F4}");
                                        softCloseAttempted = true;
                                        Log(LogLevel.Info, "Попробовали мягко закрыть MainWindow через Alt+F4.");
                                    }
                                    catch { }
                                }

                                if (!softCloseAttempted)
                                    Log(LogLevel.Warning, "Не удалось выполнить мягкое закрытие MainWindow, сразу ждём/будем убивать процесс.");

                                // --- 2) Ждём закрытия окна до минуты ---
                                int waited = 0;
                                bool closed = false;
                                while (waited < waitCloseMs)
                                {
                                    // пере-ищем окно
                                    IUIAutomationElement stillMain = null;
                                    try
                                    {
                                        var condMainAutomationId = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                        var condMainTypeWindow = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                        var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);
                                        stillMain = root.FindFirst(TreeScope.TreeScope_Descendants, condMain);
                                    }
                                    catch { }

                                    if (stillMain == null)
                                    {
                                        closed = true;
                                        Log(LogLevel.Info, "MainWindow успешно закрылось мягким способом.");
                                        break;
                                    }

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                // --- 3) Если не закрылось — убиваем процесс ---
                                if (!closed)
                                {
                                    try
                                    {
                                        int pid = (int)(mainWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ProcessIdPropertyId) ?? 0);
                                        if (pid > 0)
                                        {
                                            var proc = System.Diagnostics.Process.GetProcessById(pid);
                                            proc.Kill();
                                            Log(LogLevel.Warning, $"MainWindow не закрылось за {waitCloseMs / 1000} сек. Процесс PID={pid} был завершён принудительно.");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, "PID MainWindow не удалось получить, жёсткое завершение невозможно.");
                                        }
                                    }
                                    catch (Exception killEx)
                                    {
                                        Log(LogLevel.Error, $"Ошибка при жёстком завершении MainWindow: {killEx.Message}");
                                        throw;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при закрытии MainWindow: {ex.Message}");
                                throw;
                            }
                            #endregion

                            continue;
                        }
                        finally
                        {
                            #region Мягкое закрытие MainWindow + проверка → жёсткое завершение при необходимости
                            try
                            {
                                var automation = new CUIAutomation();
                                var root = automation.GetRootElement();

                                const int waitCloseMs = 60000; // 1 минута
                                const int pollMs = 500;

                                // --- Находим MainWindow (LanDocs) ---
                                IUIAutomationElement mainWin = null;
                                {
                                    var condMainAutomationId = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                    var condMainTypeWindow = automation.CreatePropertyCondition(
                                        UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                    var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);

                                    var mains = root.FindAll(TreeScope.TreeScope_Descendants, condMain);
                                    if (mains != null && mains.Length > 0)
                                    {
                                        for (int i = 0; i < mains.Length; i++)
                                        {
                                            var e = mains.GetElement(i);
                                            var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                                            if (name.IndexOf("LanDocs", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                mainWin = e;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (mainWin == null)
                                    throw new Exception("MainWindow (LanDocs) не найден для закрытия.");

                                // --- 1) Мягкое закрытие ---
                                bool softCloseAttempted = false;
                                try
                                {
                                    var winPatternObj = mainWin.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId);
                                    if (winPatternObj is IUIAutomationWindowPattern winPattern)
                                    {
                                        winPattern.Close();
                                        softCloseAttempted = true;
                                        Log(LogLevel.Info, "Попробовали мягко закрыть MainWindow через WindowPattern.Close().");
                                    }
                                }
                                catch
                                {
                                    // Если WindowPattern не сработал, пробуем Alt+F4
                                    try
                                    {
                                        mainWin.SetFocus();
                                        System.Windows.Forms.SendKeys.SendWait("%{F4}");
                                        softCloseAttempted = true;
                                        Log(LogLevel.Info, "Попробовали мягко закрыть MainWindow через Alt+F4.");
                                    }
                                    catch { }
                                }

                                if (!softCloseAttempted)
                                    Log(LogLevel.Warning, "Не удалось выполнить мягкое закрытие MainWindow, сразу ждём/будем убивать процесс.");

                                // --- 2) Ждём закрытия окна до минуты ---
                                int waited = 0;
                                bool closed = false;
                                while (waited < waitCloseMs)
                                {
                                    // пере-ищем окно
                                    IUIAutomationElement stillMain = null;
                                    try
                                    {
                                        var condMainAutomationId = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_AutomationIdPropertyId, "MainWindow");
                                        var condMainTypeWindow = automation.CreatePropertyCondition(
                                            UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
                                        var condMain = automation.CreateAndCondition(condMainAutomationId, condMainTypeWindow);
                                        stillMain = root.FindFirst(TreeScope.TreeScope_Descendants, condMain);
                                    }
                                    catch { }

                                    if (stillMain == null)
                                    {
                                        closed = true;
                                        Log(LogLevel.Info, "MainWindow успешно закрылось мягким способом.");
                                        break;
                                    }

                                    Thread.Sleep(pollMs);
                                    waited += pollMs;
                                }

                                // --- 3) Если не закрылось — убиваем процесс ---
                                if (!closed)
                                {
                                    try
                                    {
                                        int pid = (int)(mainWin.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ProcessIdPropertyId) ?? 0);
                                        if (pid > 0)
                                        {
                                            var proc = System.Diagnostics.Process.GetProcessById(pid);
                                            proc.Kill();
                                            Log(LogLevel.Warning, $"MainWindow не закрылось за {waitCloseMs / 1000} сек. Процесс PID={pid} был завершён принудительно.");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, "PID MainWindow не удалось получить, жёсткое завершение невозможно.");
                                        }
                                    }
                                    catch (Exception killEx)
                                    {
                                        Log(LogLevel.Error, $"Ошибка при жёстком завершении MainWindow: {killEx.Message}");
                                        throw;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при закрытии MainWindow: {ex.Message}");
                                throw;
                            }
                            #endregion
                        }
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Глобальная ошибка: {ex.Message}");
            }
            finally
            {
                Log(LogLevel.Info, "Робот завершил работу.");
            }
        }

        #region Методы

        /// <summary>
        /// Инициализация системы логирования.
        /// </summary>
        static string InitializeLogging()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDirectory))
                Directory.CreateDirectory(logDirectory);
            return logDirectory;
        }

        /// <summary>
        /// Логирование сообщений с уровнем.
        /// </summary>
        static void Log(LogLevel level, string message)
        {
            if (level > _currentLogLevel)
            {
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            string ticketFolder = GetTicketValue("ticketFolderName");
            string context = string.IsNullOrWhiteSpace(ticketFolder) ? string.Empty : $"[{ticketFolder}] ";
            string formattedMessage = $"{timestamp} [{level}] {context}{message}";

            if (!string.IsNullOrWhiteSpace(_logFilePath))
            {
                try
                {
                    File.AppendAllText(_logFilePath, formattedMessage + Environment.NewLine);
                }
                catch (IOException ex)
                {
                    Console.Error.WriteLine($"Не удалось записать сообщение в лог: {ex.Message}");
                }
            }

            Console.WriteLine(formattedMessage);
        }

        /// <summary>
        /// Загрузка параметров конфигурации.
        /// </summary>
        static bool LoadParameters(
            string filePath,
            Dictionary<string, string> targetDictionary,
            string missingFileMessage,
            string successMessage,
            string errorMessage)
        {
            if (!File.Exists(filePath))
            {
                Log(LogLevel.Error, missingFileMessage);
                return false;
            }

            try
            {
                var document = XDocument.Load(filePath);

                if (document.Root == null)
                {
                    Log(LogLevel.Error, $"Файл {filePath} не содержит корневой элемент.");
                    return false;
                }

                targetDictionary.Clear();

                foreach (var parameter in document.Root.Elements("Parameter"))
                {
                    string name = parameter.Attribute("name")?.Value;
                    string value = parameter.Attribute("value")?.Value;

                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    targetDictionary[name] = value;
                }

                Log(LogLevel.Info, successMessage);
                return true;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"{errorMessage}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Загрузка параметров конфигурации.
        /// </summary>
        static bool LoadConfig(string configPath)
        {
            if (!LoadParameters(
                    configPath,
                    _configValues,
                    "Файл config.xml не найден.",
                    "Параметры успешно загружены из config.xml",
                    "Ошибка при загрузке параметров"))
            {
                return false;
            }

            string logLevelStr = GetConfigValue("LogLevel");
            if (Enum.TryParse(logLevelStr, true, out LogLevel logLevel))
            {
                _currentLogLevel = logLevel;
                Log(LogLevel.Info, $"Уровень логирования установлен на: {_currentLogLevel}");
            }
            else if (!string.IsNullOrWhiteSpace(logLevelStr))
            {
                Log(LogLevel.Warning, $"Не удалось разобрать уровень логирования '{logLevelStr}'. Используется значение по умолчанию {_currentLogLevel}.");
            }

            return true;
        }

        /// <summary>
        /// Получение значения из параметриа конфигурации.
        /// </summary>
        static string GetConfigValue(string key) => _configValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Загрузка параметров с ППУД.
        /// </summary>
        static bool LoadConfigOrganization(string pathToOrganization)
        {
            return LoadParameters(
                pathToOrganization,
                _organizationValues,
                "Не найден файл с перечислением организаций.",
                "Список организаций успешно загружен.",
                "Ошибка при загрузке списка организаций");
        }

        /// <summary>
        /// Получение значений параметров с файла с ППУД.
        /// </summary>
        static string GetConfigOrganization(string key) => _organizationValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Получение значения из текущей заявки.
        /// </summary>
        static string GetTicketValue(string key) => _ticketValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Метод очистки логов
        /// </summary>
        static void CleanOldLogs(string logDirectory, int retentionDays)
        {
            foreach (var log in Directory.EnumerateFiles(logDirectory, "*.txt").Where(f => File.GetCreationTime(f) < DateTime.Now.AddDays(-retentionDays)))
            {
                try
                {
                    File.Delete(log);
                    Log(LogLevel.Info, $"Лог-файл {log} удален");
                }
                catch (Exception e)
                {
                    Log(LogLevel.Error, $"Ошибка при удалении файла лога {log}: {e.Message}");
                }
            }
        }

        /// <summary>
        /// Получение массива с файлами и папками
        /// </summary>
        static string[] GetFilesAndFoldersFromDirectory(string folder)
        {
            try
            {
                return Directory.GetFiles(folder)
                    .Concat(Directory.GetDirectories(folder))
                    .ToArray();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при получении файлов и папок из папки {folder}: {ex.Message}");
                return Array.Empty<string>();  // Возвращаем пустой массив при ошибке
            }
        }

        /// <summary>
        /// Поиск по наименованию папки.
        /// </summary>
        static string GetFoldersSearchDirectory(string folder, string dirName)
        {
            try
            {
                return Directory.GetDirectories(folder, dirName, SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при поиске папки {dirName} в {folder}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получение всех файлов в папке.
        /// </summary>
        static string[] GetFileshDirectory(string folder)
        {
            try
            {
                return Directory.GetFiles(folder);
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка получении файлов в папке {folder}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Поиск файла по названию.
        /// </summary>
        static string GetFileSearchDirectory(string directory, string searchPattern)
        {
            try
            {
                return Directory.GetFiles(directory, searchPattern, SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при поиске файлов в папке {directory}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Проверяет и создает указанные директории в baseFolder папке. Возвращает false при ошибке создания.
        /// </summary>
        static bool EnsureDirectoriesExist(string baseFolder, params string[] folderNames)
        {
            foreach (var folderName in folderNames)
            {
                string folderPath = Path.Combine(baseFolder, folderName);
                if (!CreateDirectoryWithLogging(folderPath))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Создает директорию по указанному пути, если она не существует, и логирует результат.
        /// </summary>
        static bool CreateDirectoryWithLogging(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    Log(LogLevel.Debug, $"Папка {path} успешно создана.");
                }
                return true;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Не удалось создать папку {path}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Создает папки для различных типов файлов и перемещает файлы в соответствующие папки.
        /// Если файл не является .xlsx, .pdf или .zip, он перемещается в папку "error".
        /// </summary>
        static FolderPaths CreateFolderMoveFiles(string creatingFolder, string[] arrayFiles)
        {
            // Папки для разных типов файлов
            var folderPaths = new FolderPaths
            {
                XlsxFolder = Path.Combine(creatingFolder, "xlsx"),
                PdfFolder = Path.Combine(creatingFolder, "pdf"),
                ZipFolder = Path.Combine(creatingFolder, "zip"),
                ErrorFolder = Path.Combine(creatingFolder, "error"),
                DocumentFolder = Path.Combine(creatingFolder, "document")
            };

            foreach (var file in arrayFiles)
            {
                try
                {
                    // Пропускаем папки (только файлы)
                    if (!File.Exists(file))
                    {
                        continue; // Это папка, пропускаем
                    }

                    string extension = Path.GetExtension(file).ToLower();
                    string destinationFolder = GetDestinationFolder(extension, folderPaths);
                    string destination = Path.Combine(destinationFolder, Path.GetFileName(file));

                    // Перемещаем файл
                    File.Move(file, destination);

                    // Логируем результат
                    if (extension == ".xlsx" || extension == ".pdf" || extension == ".zip")
                    {
                        Log(LogLevel.Debug, $"Перемещен файл {file} в {destinationFolder}");
                    }
                    else
                    {
                        Log(LogLevel.Warning, $"Файл {file} не является .xlsx, .pdf или .zip, перемещен в папку error.");
                    }
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Error, $"Ошибка при перемещении файла {file}: {ex.Message}");
                }
            }
            return folderPaths;
        }

        /// <summary>
        /// Определяет папку назначения для файла в зависимости от его расширения.
        /// </summary>
        static string GetDestinationFolder(string extension, FolderPaths folderPaths)
        {
            string destinationFolder;

            switch (extension)
            {
                case ".xlsx":
                    destinationFolder = folderPaths.XlsxFolder;
                    break;
                case ".pdf":
                    destinationFolder = folderPaths.PdfFolder;
                    break;
                case ".zip":
                    destinationFolder = folderPaths.ZipFolder;
                    break;
                default:
                    destinationFolder = folderPaths.ErrorFolder;
                    break;
            }
            return destinationFolder;
        }

        /// <summary>
        /// Класс, представляющий пути к различным папкам для хранения файлов.
        /// </summary>
        public class FolderPaths
        {
            public string XlsxFolder { get; set; }
            public string PdfFolder { get; set; }
            public string ZipFolder { get; set; }
            public string ErrorFolder { get; set; }
            public string DocumentFolder { get; set; }
        }

        /// <summary>
        /// Метод парсинга Json файла заявки
        /// </summary>
        public static (string OrgTitle, string Title, string FormType, string ppudOrganization) ParseJsonFile(string filePath)
        {
            // Проверка существования файла
            if (!File.Exists(filePath))
            {
                Log(LogLevel.Fatal, $"Файл не найден: {filePath}");
                throw new FileNotFoundException($"Файл не найден: {filePath}");
            }

            Log(LogLevel.Debug, $"Начинается обработка файла: {filePath}");

            // Чтение содержимого файла
            string jsonContent;
            try
            {
                jsonContent = File.ReadAllText(filePath);
                Log(LogLevel.Debug, $"Файл успешно прочитан: {filePath}");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Ошибка чтения файла {filePath}: {ex.Message}");
                throw new IOException($"Ошибка чтения файла: {ex.Message}");
            }

            // Проверка на пустое содержимое
            if (string.IsNullOrWhiteSpace(jsonContent))
            {
                Log(LogLevel.Fatal, $"Файл пуст или содержит только пробелы: {filePath}");
                throw new InvalidOperationException("Файл пуст или содержит только пробелы.");
            }

            // Парсинг JSON
            try
            {
                JToken jsonToken = JToken.Parse(jsonContent);

                JObject jsonObject;
                if (jsonToken is JObject obj)
                {
                    jsonObject = obj;
                }
                else if (jsonToken is JArray array && array.Count > 0 && array[0] is JObject firstObj)
                {
                    jsonObject = firstObj; // Если JSON - массив, берем первый объект
                }
                else
                {
                    Log(LogLevel.Fatal, $"Неверный формат JSON: ожидался объект или массив объектов в файле {filePath}. Проверьте файл заявки.");
                    throw new InvalidOperationException("Неверный формат JSON: ожидался объект или массив объектов. Проверьте файл заявки.");
                }

                /// Извлечение значений с логированием ошибок
                string orgTitle = jsonObject?["orgFil"]?["title"]?.ToString();
                if (string.IsNullOrEmpty(orgTitle))
                {
                    Log(LogLevel.Fatal, $"Поле 'orgFil.title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'orgFil.title' отсутствует или пустое.");
                }

                string title = jsonObject?["title"]?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    Log(LogLevel.Fatal, $"Поле 'title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'title' отсутствует или пустое.");
                }

                string formType = jsonObject?["formTypeInt"]?["title"]?.ToString()?.Trim();
                if (string.IsNullOrEmpty(formType))
                {
                    Log(LogLevel.Fatal, $"Поле 'formTypeInt.title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'formTypeInt.title' отсутствует или пустое.");
                }

                // Пытаемся найти организацию по названию
                var matchingKeyValue = _organizationValues.FirstOrDefault(kv => kv.Key == orgTitle);
                if (matchingKeyValue.Key == null)
                {
                    Log(LogLevel.Fatal, $"ППУД для организации [{orgTitle}] не найдена в коллекции _organizationValues. JSON: {filePath}");
                    throw new InvalidOperationException($"ППУД с ключом '{orgTitle}' не найдена.");
                }

                string ppudOrganization = matchingKeyValue.Value;

                return (orgTitle, title, formType, ppudOrganization);
            }
            catch (JsonReaderException ex)
            {
                Log(LogLevel.Error, $"Ошибка парсинга JSON в файле {filePath}: {ex.Message}");
                throw new InvalidOperationException($"Ошибка парсинга JSON: {ex.Message}");
            }
        }
        static string ReplaceIgnoreCase(string input, string search, string replacement)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(search))
                return input;

            var sb = new System.Text.StringBuilder(input.Length);
            int i = 0;
            while (true)
            {
                int pos = input.IndexOf(search, i, StringComparison.OrdinalIgnoreCase);
                if (pos < 0)
                {
                    sb.Append(input, i, input.Length - i);
                    break;
                }
                sb.Append(input, i, pos - i);
                sb.Append(replacement);
                i = pos + search.Length;
            }
            return sb.ToString();
        }
        static string NormalizeBaseName(string nameWithoutExt)
        {
            if (string.IsNullOrWhiteSpace(nameWithoutExt)) return string.Empty;
            string s = nameWithoutExt.Trim();
            s = ReplaceIgnoreCase(s, "ОЦО", "").Trim();

            if (s.EndsWith("OK", StringComparison.OrdinalIgnoreCase) ||
                s.EndsWith("ОК", StringComparison.OrdinalIgnoreCase))
                s = s.Substring(0, s.Length - 2).Trim();

            return s.Replace(" ", string.Empty).ToLowerInvariant();
        }

        // === helper: строка заканчивается на OK/ОК ===
        static bool EndsWithOkSuffix(string nameWithoutExt)
        {
            return nameWithoutExt.EndsWith("OK", StringComparison.OrdinalIgnoreCase) ||
                   nameWithoutExt.EndsWith("ОК", StringComparison.OrdinalIgnoreCase);
        }

        // === helper: безопасно добавить суффикс " ОК" перед расширением, с разруливанием коллизий ===
        static string TryAppendCyrillicOkSuffix(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            if (dir == null)
                throw new InvalidOperationException($"Не удалось определить папку для файла: {filePath}");
            string nameNoExt = Path.GetFileNameWithoutExtension(filePath);
            string ext = Path.GetExtension(filePath);

            if (EndsWithOkSuffix(nameNoExt)) return filePath; // уже помечен

            string candidate = Path.Combine(dir, nameNoExt + " ОК" + ext);
            if (!File.Exists(candidate)) return candidate;

            // если занято — добавляем (2), (3), ...
            int i = 2;
            while (true)
            {
                string c = Path.Combine(dir, $"{nameNoExt} ОК ({i}){ext}");
                if (!File.Exists(c)) return c;
                i++;
            }
        }

        // === XlsxContainsPDF: без переименований, только выбор уникальных xlsx без pdf ===
        static string[] XlsxContainsPDF(string xlsxFolder, string pdfFolder)
        {
            var existingPdfNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (var pdfPath in Directory.GetFiles(pdfFolder, "*.pdf"))
            {
                var pdfBase = Path.GetFileNameWithoutExtension(pdfPath);
                var norm = NormalizeBaseName(pdfBase);
                if (!string.IsNullOrEmpty(norm)) existingPdfNames.Add(norm);
            }

            var xlsxFiles = Directory.GetFiles(xlsxFolder, "*.xlsx")
                                     .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                     .ToArray();

            var seenXlsx = new HashSet<string>(StringComparer.Ordinal);
            var result = new List<string>();

            foreach (var xlsx in xlsxFiles)
            {
                string nameNoExt = Path.GetFileNameWithoutExtension(xlsx)?.Trim() ?? "";
                string norm = NormalizeBaseName(nameNoExt);
                if (string.IsNullOrEmpty(norm))
                {
                    Log(LogLevel.Warning, $"[!] Пустое нормализованное имя для [{xlsx}] — пропуск.");
                    continue;
                }

                // уникальность по смыслу
                if (!seenXlsx.Add(norm))
                {
                    Log(LogLevel.Debug, $"[=] Пропуск дубликата по имени: [{nameNoExt}]");
                    continue;
                }

                bool hasPdf = existingPdfNames.Contains(norm);
                if (!hasPdf)
                {
                    result.Add(xlsx);
                    Log(LogLevel.Warning, $"[-] Для файла [{nameNoExt}] PDF не найден — добавлен в очередь на конвертацию.");
                }
                else
                {
                    Log(LogLevel.Debug, $"[+] Для файла [{nameNoExt}] PDF уже существует — конвертация не требуется.");
                }
            }

            return result.ToArray();
        }

        // === ConvertToPdf: после успешного экспорта — ПЕРЕИМЕНОВАТЬ XLSX, добавив " ОК" ===
        static void ConvertToPdf(IEnumerable<string> xlsxFiles, string outputFolder)
        {
            Excel.Application excelApplication = null;

            try
            {
                excelApplication = new Excel.Application { Visible = false, DisplayAlerts = false };

                foreach (var xlsxPath in xlsxFiles)
                {
                    Excel.Workbook workbook = null;
                    bool exported = false;
                    string renameFrom = xlsxPath;
                    string renameTo = null;

                    try
                    {
                        Log(LogLevel.Debug, $"Обработка файла: {xlsxPath}");

                        string nameNoExt = Path.GetFileNameWithoutExtension(xlsxPath);
                        string pdfBaseName = EndsWithOkSuffix(nameNoExt)
                            ? nameNoExt.Substring(0, nameNoExt.Length - 2).Trim()
                            : nameNoExt;

                        string sanitized = string.Join("_", pdfBaseName.Split(Path.GetInvalidFileNameChars()));
                        string outputFile = Path.Combine(outputFolder, $"{sanitized}.pdf");
                        Log(LogLevel.Debug, $"Выходной файл: {outputFile}");

                        if (File.Exists(outputFile))
                        {
                            Log(LogLevel.Info, $"[=] PDF уже существует, пропуск конвертации: {outputFile}");
                            continue;
                        }

                        // Открываем в read-only; всё равно будет блокировка, но не модифицируем файл
                        workbook = excelApplication.Workbooks.Open(Filename: xlsxPath, ReadOnly: true, IgnoreReadOnlyRecommended: true);
                        workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile);
                        exported = true;

                        // целевое имя для переименования после закрытия книги
                        renameTo = TryAppendCyrillicOkSuffix(renameFrom);

                        Log(LogLevel.Debug, $"[✓] Успешно конвертировано: {outputFile}");
                    }
                    catch (Exception ex)
                    {
                        Log(LogLevel.Error, $"Ошибка обработки файла '{xlsxPath}': {ex.Message}");
                    }
                    finally
                    {
                        // 1) Закрываем книгу и освобождаем COM
                        if (workbook != null)
                        {
                            try { workbook.Close(false); } catch { }
                            try { Marshal.ReleaseComObject(workbook); } catch { }
                            workbook = null;
                        }

                        // 2) Принудительно соберём мусор, чтобы быстрее отпустились COM-хэндлы
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        // 3) Если экспорт удался — ПЕРЕИМЕНОВАТЬ исходник с ретраями
                        if (exported && renameTo != null && !renameFrom.Equals(renameTo, StringComparison.OrdinalIgnoreCase))
                        {
                            TryRenameWithRetries(renameFrom, renameTo, attempts: 10, delayMs: 300);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при работе с Excel: {ex.Message}");
            }
            finally
            {
                if (excelApplication != null)
                {
                    try { excelApplication.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(excelApplication); } catch { }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Метод нахождения и удаления процесса приложения
        /// </summary>
        static void KillExcelProcesses(string NameProceses)
        {
            try
            {
                string currentUser = Environment.UserName; // Получение имени текущего пользователя

                foreach (var process in Process.GetProcessesByName(NameProceses))
                {
                    try
                    {
                        if (IsProcessOwnedByCurrentUser(process))
                        {
                            Log(LogLevel.Debug, $"Завершаем процесс {NameProceses} с ID {process.Id}, пользователь: {currentUser}");
                            process.Kill();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log(LogLevel.Error, $"Ошибка при завершении процесса {NameProceses} с ID {process.Id}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при завершении процессов {NameProceses}: {ex.Message}");
            }
        }

        /// <summary>
        /// Метод нахождения процесса приложения по имени у текущей УЗ
        /// </summary>
        static bool IsProcessOwnedByCurrentUser(Process process)
        {
            try
            {
                // Проверка владельца процесса через WMI
                var query = $"SELECT * FROM Win32_Process WHERE ProcessId = {process.Id}";
                using (var searcher = new ManagementObjectSearcher(query))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var outParams = obj.InvokeMethod("GetOwner", null, null);
                        if (outParams != null && outParams.Properties["User"] != null)
                        {
                            string user = outParams.Properties["User"].Value.ToString();
                            return string.Equals(user, Environment.UserName, StringComparison.OrdinalIgnoreCase);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при определении владельца процесса {process.Id}: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Метод перемещения профиля landocs
        /// </summary>
        static void MoveCustomProfileLandocs(string customFile, string landocsProfileFolder)
        {
            try
            {
                // Проверяем существование исходного файла
                if (!File.Exists(customFile))
                {
                    throw new FileNotFoundException($"Ошибка: исходный файл профиля landocs '{customFile}' не найден.");
                }

                // Убедимся, что папка назначения существует
                if (!Directory.Exists(landocsProfileFolder))
                {
                    throw new FileNotFoundException($"Ошибка: папка с профилями landocs '{customFile}' не найден.");
                }

                // Формируем полный путь к файлу в папке назначения
                string destinationFilePath = Path.Combine(landocsProfileFolder, Path.GetFileName(customFile));

                // Если файл назначения существует, меняем его расширение на .bak
                if (File.Exists(destinationFilePath))
                {
                    string backupFilePath = Path.ChangeExtension(destinationFilePath, ".bak");

                    // Удаляем старый .bak файл, если он существует
                    if (File.Exists(backupFilePath))
                    {
                        File.Delete(backupFilePath);
                    }

                    File.Move(destinationFilePath, backupFilePath);
                    Log(LogLevel.Debug, $"Выполнил резервную копию файла профиля [{destinationFilePath}] переименован в [{backupFilePath}].");
                }

                // Перемещаем новый файл
                File.Copy(customFile, destinationFilePath);

                Log(LogLevel.Debug, $"Кастомный файл профиля landocs успешно перемещен из '{customFile}' в '{destinationFilePath}'.");
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                Log(LogLevel.Fatal, $"Ошибка перемещения профиля: {ex.Message}");

                // Бросаем исключение, чтобы завершить работу приложения
                throw new ApplicationException($"Критическая ошибка: {ex.Message}", ex);
            }
        }

        static void OpenLandocsAndNavigateToMessages()
        {
            try
            {
                Log(LogLevel.Info, "Запускаю Landocs для перехода во вкладку [Сообщения].");

                string customFile = GetConfigValue("ConfigLandocsCustomFile");
                string landocsProfileFolder = GetConfigValue("ConfigLandocsFolder");
                MoveCustomProfileLandocs(customFile, landocsProfileFolder);

                string appLandocsPath = GetConfigValue("AppLandocsPath");
                var appElement = LaunchAndFindWindow(appLandocsPath, "_robin_landocs (Мой LanDocs) - Избранное - LanDocs", 300);

                if (appElement == null)
                {
                    throw new Exception("Окно Landocs не найдено при переходе во вкладку [Сообщения].");
                }

                Thread.Sleep(5000);

                string xpathHomeTab = "Pane[3]/Tab/TabItem[1]";
                var homeTab = FindElementByXPath(appElement, xpathHomeTab, 60);
                if (homeTab != null)
                {
                    ClickElementWithMouse(homeTab);
                    Log(LogLevel.Info, "Открыта вкладка [Главная] перед переходом в [Сообщения].");
                }
                else
                {
                    Log(LogLevel.Warning, "Не удалось найти вкладку [Главная] перед переходом в [Сообщения].");
                }

                string xpathMessagesButton = "Pane[1]/Pane/Pane[1]/Pane/Pane/Button[1]";
                var messagesButton = FindElementByXPath(appElement, xpathMessagesButton, 60);
                if (messagesButton == null)
                {
                    throw new Exception("Элемент [Сообщения] в навигационном меню не найден.");
                }

                if (!TryInvokeElement(messagesButton))
                {
                    ClickElementWithMouse(messagesButton);
                }

                Log(LogLevel.Info, "Вкладка [Сообщения] успешно открыта.");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при переходе во вкладку [Сообщения]: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Метод запуска landocs
        /// </summary>
        public static IUIAutomationElement LaunchAndFindWindow(string appPath, string windowName, int maxWaitTime)
        {
            try
            {
                var automation = new CUIAutomation();
                var rootElement = automation.GetRootElement();

                Log(LogLevel.Info, $"Запуск приложения: {appPath}");
                var appProcess = Process.Start(appPath);

                if (appProcess == null)
                {
                    Log(LogLevel.Error, "Не удалось запустить приложение.");
                    throw new ApplicationException("Критическая ошибка: Не удалось запустить приложение.");
                }

                IUIAutomationElement appElement = null;
                int elapsedSeconds = 0;

                Log(LogLevel.Info, $"Поиск окна приложения с именем: [{windowName}]. Время ожидания:[{maxWaitTime}] сек.");

                while (elapsedSeconds < maxWaitTime && appElement == null)
                {
                    IUIAutomationCondition condition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, windowName);
                    appElement = rootElement.FindFirst(TreeScope.TreeScope_Children, condition);

                    if (appElement == null)
                    {
                        Thread.Sleep(1000);
                        elapsedSeconds++;

                        Log(LogLevel.Debug, $"Ожидание окна приложения: [{windowName}]. Прошло [{elapsedSeconds}] секунд...");

                        // Каждые 10 секунд - лог уровня Info
                        if (elapsedSeconds % 10 == 0)
                        {
                            Log(LogLevel.Warning, $"Ожидание окна приложения: [{windowName}]. Прошло [{elapsedSeconds}] секунд.");
                        }
                    }
                }

                if (appElement != null)
                {
                    Log(LogLevel.Info, "Landocs успешно запустился.");
                }
                else
                {
                    Log(LogLevel.Error, "Окно приложения не найдено после максимального времени ожидания.");
                    throw new ApplicationException($"Критическая ошибка: Окно приложения '{windowName}' не найдено.");
                }

                return appElement;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Ошибка при запуске или поиске окна приложения: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Метод клика на элемент (Эмуляция программного нажатия)
        /// </summary>
        /*static bool TryInvokeElement(IUIAutomationElement element, int attempts = 3, int retryDelayMs = 300, int staTimeoutMs = 5000)
        {
            if (element == null) return false;

            // Быстрые проверки доступности
            try
            {
                if (element.CurrentIsEnabled == 0)
                {
                    Console.WriteLine("Элемент недоступен для взаимодействия (IsEnabled=0).");
                    return false;
                }
            }
            catch (COMException) { *//* элемент мог устареть — продолжим по общему сценарию с ретраями *//* }

            for (int tryIndex = 1; tryIndex <= attempts; tryIndex++)
            {
                try
                {
                    // Держим всю операцию в STA
                    Exception staEx = null;
                    bool ok = false;

                    var t = new Thread(() =>
                    {
                        try
                        {
                            // На всякий случай фокус
                            try { element.SetFocus(); } catch { *//* не критично *//* }

                            // 1) InvokePattern
                            if (element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) is IUIAutomationInvokePattern invPat)
                            {
                                invPat.Invoke();
                                ok = true;
                                return;
                            }

                            // 2) LegacyIAccessible (если доступен)
                            if (element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) is IUIAutomationLegacyIAccessiblePattern acc)
                            {
                                acc.DoDefaultAction();
                                ok = true;
                                return;
                            }

                            // 3) Фолбэк — клик мышью по элементу (используйте ваш метод)
                            // Если у вас уже есть ClickElementWithMouse(element) — вызовите его:
                            ClickElementWithMouse(element);
                            ok = true;
                        }
                        catch (Exception ex)
                        {
                            staEx = ex;
                        }
                    });

                    t.IsBackground = true;
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();

                    if (!t.Join(staTimeoutMs))
                    {
                        try { t.Interrupt(); } catch { }
                        throw new TimeoutException("Invoke не завершился в отведённое время STA.");
                    }

                    if (staEx != null) throw staEx;

                    if (ok)
                    {
                        Console.WriteLine("Действие выполнено (Invoke/LegacyIAccessible/Mouse).");
                        return true;
                    }

                    // если сюда дошли — считаем как неуспех попытки
                    throw new InvalidOperationException("Не удалось выполнить действие для элемента.");
                }
                catch (COMException ex) when ((uint)ex.HResult == 0x80040200) // элемент устарел/недоступен
                {
                    Console.WriteLine($"COM 0x80040200 на попытке {tryIndex}: {ex.Message}. Повтор через {retryDelayMs} мс.");
                    Thread.Sleep(retryDelayMs);
                    // Полезно перед повтором заново отыскать элемент, если у вас есть локатор (AutomationId/Name).
                    // Если локатора нет — следующий заход может сработать, когда провайдер «проснётся».
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Попытка {tryIndex} завершилась ошибкой: {ex.Message}");
                    Thread.Sleep(retryDelayMs);
                }
            }

            Console.WriteLine("Не удалось выполнить действие: исчерпаны попытки.");
            return false;
        }*/

        /// <summary>
        /// Метод клика на элемент (Эмуляция физического нажатия)
        /// </summary>
        static void ClickElementWithMouse(IUIAutomationElement element)
        {
            try
            {
                // Получение границ элемента
                object boundingRectValue = element.GetCurrentPropertyValue(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);

                // Проверяем, что значение границ корректное
                if (!(boundingRectValue is double[] boundingRectangle) || boundingRectangle.Length != 4)
                {
                    Log(LogLevel.Warning, "Не удалось получить или обработать границы элемента.");
                    throw new InvalidOperationException("Некорректные границы элемента.");
                }

                // Извлечение координат
                int left = (int)boundingRectangle[0];
                int top = (int)boundingRectangle[1];
                int right = (int)boundingRectangle[2];
                int bottom = (int)boundingRectangle[3];

                // Проверяем, что размеры валидны
                /*if (right <= left || bottom <= top)
                {
                    Log(LogLevel.Warning, "Границы элемента некорректны.");
                    throw new InvalidOperationException("Неверные размеры элемента.");
                }*/

                // Расчет центра элемента
                int x = left + right / 2;
                int y = top + bottom / 2;

                // Устанавливаем курсор на центр элемента
                if (!SetCursorPos(x, y))
                {
                    Log(LogLevel.Error, $"Не удалось установить курсор на позицию: X={x}, Y={y}");
                    throw new InvalidOperationException("Ошибка установки позиции курсора.");
                }

                // Небольшая задержка перед кликом
                Thread.Sleep(100);

                // Выполняем клик
                mouse_event((int)MouseFlags.LeftDown, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(200);
                mouse_event((int)MouseFlags.LeftUp, 0, 0, 0, UIntPtr.Zero);

                Log(LogLevel.Info, $"Клик выполнен по элементу в центре: X={x}, Y={y}");
            }
            catch (COMException ex)
            {
                Log(LogLevel.Error, $"COM-ошибка при попытке кликнуть по элементу: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Общая ошибка при клике по элементу: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Метод поиска элемента по xpath
        /// </summary>
        static IUIAutomationElement FindElementByXPath(IUIAutomationElement root, string xpath, int secondsToWait)
        {
            var automation = new CUIAutomation();
            IUIAutomationCondition trueCondition = automation.CreateTrueCondition();
            string[] parts = xpath.Split('/');
            IUIAutomationElement currentElement = root;

            int elapsedSeconds = 0;
            const int checkInterval = 500;

            while (elapsedSeconds < secondsToWait)
            {
                foreach (var part in parts)
                {
                    if (currentElement == null)
                    {
                        Console.WriteLine("Текущий элемент равен null, поиск прерван.");
                        return null;
                    }

                    var split = part.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                    string type = split[0];
                    int index = split.Length > 1 ? int.Parse(split[1]) - 1 : 0;

                    // Проверяем, что мы можем найти дочерние элементы
                    IUIAutomationElementArray children = currentElement.FindAll(TreeScope.TreeScope_Children, trueCondition);

                    if (children == null || children.Length == 0)
                    {
                        Console.WriteLine("Дочерние элементы не найдены.");
                        return null;
                    }

                    bool found = false;
                    int typeCount = 0;

                    for (int i = 0; i < children.Length; i++)
                    {
                        IUIAutomationElement child = children.GetElement(i);

                        if (child != null && child.CurrentControlType == GetControlType(type))
                        {
                            if (typeCount == index)
                            {
                                currentElement = child;
                                found = true;
                                break;
                            }
                            typeCount++;
                        }
                    }

                    if (!found)
                    {
                        currentElement = null;
                        break;
                    }
                }

                if (currentElement != null)
                {
                    return currentElement;
                }

                Thread.Sleep(checkInterval);
                elapsedSeconds += checkInterval / 1000;
            }

            return null;
        }

        /// <summary>
        /// Метод поиска элемента по параметру Name
        /// </summary>
        static IUIAutomationElement FindElementByName(IUIAutomationElement root, string name, int secondsToWait, int pollMs = 500)
        {
            var automation = new CUIAutomation();
            var nameCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name);

            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < TimeSpan.FromSeconds(secondsToWait))
            {
                try
                {
                    var element = root.FindFirst(TreeScope.TreeScope_Descendants, nameCondition);
                    if (element != null)
                        return element;
                }
                catch { /* UIA может кидать исключения, если дерево меняется — игнорируем и продолжаем */ }

                Thread.Sleep(pollMs);
            }
            return null;
        }

        static IUIAutomationElement FindElementByNameContains(IUIAutomationElement root, string substring, int secondsToWait, int pollMs = 500)
        {
            var automation = new CUIAutomation();
            var textType = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_TextControlTypeId);

            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < TimeSpan.FromSeconds(secondsToWait))
            {
                try
                {
                    var coll = root.FindAll(TreeScope.TreeScope_Descendants, textType);
                    for (int i = 0; i < coll?.Length; i++)
                    {
                        var e = coll.GetElement(i);
                        var name = (e.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string) ?? string.Empty;
                        if (name.IndexOf(substring, StringComparison.OrdinalIgnoreCase) >= 0)
                            return e;
                    }
                }
                catch { }
                Thread.Sleep(pollMs);
            }
            return null;
        }

        /// <summary>
        /// Метод возвращающий тип ControlType
        /// </summary>
        static int GetControlType(string type)
        {
            type = type.ToLower();

            switch (type)
            {
                case "pane": return UIA_ControlTypeIds.UIA_PaneControlTypeId;
                case "table": return UIA_ControlTypeIds.UIA_TableControlTypeId;
                case "tab": return UIA_ControlTypeIds.UIA_TabControlTypeId;
                case "tabitem": return UIA_ControlTypeIds.UIA_TabItemControlTypeId;
                case "button": return UIA_ControlTypeIds.UIA_ButtonControlTypeId;
                case "group": return UIA_ControlTypeIds.UIA_GroupControlTypeId;
                case "checkbox": return UIA_ControlTypeIds.UIA_CheckBoxControlTypeId;
                case "combobox": return UIA_ControlTypeIds.UIA_ComboBoxControlTypeId;
                case "edit": return UIA_ControlTypeIds.UIA_EditControlTypeId;
                case "text": return UIA_ControlTypeIds.UIA_TextControlTypeId;
                case "window": return UIA_ControlTypeIds.UIA_WindowControlTypeId;
                case "custom": return UIA_ControlTypeIds.UIA_CustomControlTypeId;
                case "tree": return UIA_ControlTypeIds.UIA_TreeControlTypeId;
                case "toolbar": return UIA_ControlTypeIds.UIA_ToolBarControlTypeId;
                case "dataitem": return UIA_ControlTypeIds.UIA_DataItemControlTypeId;
                default: return UIA_ControlTypeIds.UIA_PaneControlTypeId;
            }
        }

        /// <summary>
        /// Метод возвращающий элемент на котором сейчас установлен фокус
        /// </summary>
        static IUIAutomationElement GetFocusedElement()
        {
            var automation = new CUIAutomation();
            IUIAutomationElement focusedElement = automation.GetFocusedElement();

            if (focusedElement != null)
            {
                try
                {
                    Console.WriteLine("Элемент с фокусом найден:");
                    Console.WriteLine($"Имя элемента: {focusedElement.CurrentName}");
                    Console.WriteLine($"Тип элемента: {focusedElement.CurrentControlType}");
                    Console.WriteLine($"Тип элемента: {focusedElement.CurrentLocalizedControlType}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при получении информации об элементе с фокусом: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Элемент с фокусом не найден.");
            }

            return focusedElement;
        }

        /// <summary>
        /// Метод возвращающий элемент окна с ошибкой
        /// </summary>
        static IUIAutomationElement GetErrorWindowElement(IUIAutomationElement rootElement, string echildrenNameWindow)
        {
            var targetWindowError = FindElementByName(rootElement, echildrenNameWindow, 60);

            // Проверяем значение свойства Name элемента
            if (targetWindowError != null)
            {
                // Создаем условия для поиска title и message
                var automation = new CUIAutomation();

                // Условие для поиска элемента сообщения (message)
                var messageCondition = automation.CreatePropertyCondition(
                    UIA_PropertyIds.UIA_ControlTypePropertyId,
                    UIA_ControlTypeIds.UIA_TextControlTypeId
                );
                var messageElement = targetWindowError.FindFirst(TreeScope.TreeScope_Children, messageCondition);

                string message = messageElement != null
                    ? messageElement.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string
                    : "Сообщение не найдено";
                Log(LogLevel.Fatal, $"Появилось окно [Ошибка], текст сообщения: [{message}]");
                // Ищем кнопку "ОК"
                var buttonOk = FindElementByName(targetWindowError, "&ОК", 60);

                throw new Exception("Появилось окно ошибки. Работа робота завершена.");
            }
            else
            {
                throw new Exception($"Появилось окно ошибки. Не удалось определить элемент. Робот завершает работу.");
            }
        }

        /// <summary>
        /// Метод возвращающий ключ контрагента найденного по ИНН и КПП
        /// </summary>
        static int? FindCounterpartyKey(Dictionary<int, string[]> counterpartyElements, string innValue, string kppValue, string counterpartyName = null)
        {
            // Приводим значения ИНН и КПП к единому формату заранее
            string formattedInnValue = $"ИНН:{innValue}".Replace(" ", "").Trim().ToLower();
            string formattedKppValue = string.IsNullOrEmpty(kppValue) ? null : $"КПП:{kppValue}".Replace(" ", "").Trim().ToLower();
            string formattedCounterpartyName = string.IsNullOrEmpty(counterpartyName) ? null : counterpartyName.Replace(" ", "").Trim().ToLower();

            foreach (var kvp in counterpartyElements)
            {
                // Очищаем элементы списка контрагентов от лишних пробелов и приводим к нижнему регистру один раз
                var formattedElements = kvp.Value.Select(x => x.Replace(" ", "").Trim().ToLower()).ToList();

                // Проверяем наличие ИНН
                bool innMatch = formattedElements.Contains(formattedInnValue);

                // Проверяем наличие КПП (если оно задано)
                bool kppMatch = string.IsNullOrEmpty(formattedKppValue) || formattedElements.Contains(formattedKppValue);

                // Если КПП отсутствует, проверяем по имени контрагента
                bool nameMatch = string.IsNullOrEmpty(formattedKppValue) && !string.IsNullOrEmpty(formattedCounterpartyName) &&
                                 formattedElements.Any(x => x.Contains(formattedCounterpartyName));

                // Если найдено совпадение по ИНН и либо КПП, либо имени
                if (innMatch && (kppMatch || nameMatch))
                {
                    return kvp.Key;
                }
            }
            return null; // Возвращаем null, если совпадений не найдено
        }

        /// <summary>
        /// Метод возвращающий параметры с названия файла для landocs
        /// </summary>
        static FileData GetParseNameFile(string fileName)
        {
            // Регулярное выражение для парсинга строки
            var match = Regex.Match(fileName,
                @"Акт св П \d+\s+(.*?)\s+№(\S+)\s+(\d{2}\.\d{2}\.\d{2})_(\d+)_?(\d+)?");

            if (match.Success)
            {
                return new FileData
                {
                    CounterpartyName = match.Groups[1].Value.Trim(),
                    Number = match.Groups[2].Value.Trim(),
                    FileDate = match.Groups[3].Value.Trim(),
                    INN = match.Groups[4].Value.Trim(),
                    KPP = match.Groups[5].Success ? match.Groups[5].Value.Trim() : null
                };
            }
            else
            {
                Console.WriteLine($"Не удалось распознать файл: {fileName}");
                //Добавить перемещение в папку error
                return null;
            }
        }

        /// <summary>
        /// Класс, с параметрами файла для landocs
        /// </summary>
        public class FileData
        {
            public string CounterpartyName { get; set; }
            public string Number { get; set; }
            public string FileDate { get; set; }
            public string INN { get; set; }
            public string KPP { get; set; }
        }
        #endregion

        /// <summary>
        /// Переключает раскладку клавиатуры на английскую (en-US), если она доступна в системе.
        /// </summary>
        public static void EnsureEnglishKeyboardLayout()
        {
            try
            {
                var englishCulture = new CultureInfo("en-US");
                var englishInput = InputLanguage.InstalledInputLanguages
                    .Cast<InputLanguage>()
                    .FirstOrDefault(lang => string.Equals(lang.Culture.Name, englishCulture.Name, StringComparison.OrdinalIgnoreCase));

                if (englishInput == null)
                {
                    Log(LogLevel.Warning, "Английская раскладка клавиатуры (en-US) не установлена. Переключение не выполнено.");
                    return;
                }

                var currentInput = InputLanguage.CurrentInputLanguage;
                if (!string.Equals(currentInput?.Culture.Name, englishCulture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    InputLanguage.CurrentInputLanguage = englishInput;
                    Log(LogLevel.Info, "Раскладка клавиатуры переключена на английскую (en-US).");
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Warning, $"Не удалось переключить раскладку клавиатуры на английскую: {ex.Message}");
            }
        }

        static bool TryInvokeElement(IUIAutomationElement element, int attempts = 3, int retryDelayMs = 250)
        {
            if (element == null) return false;

            for (int i = 1; i <= attempts; i++)
            {
                try
                {
                    // Фокус (не критично)
                    try { element.SetFocus(); } catch { }

                    // Если есть Invoke — пробуем
                    var pat = element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) as IUIAutomationInvokePattern;
                    if (pat != null)
                    {
                        try
                        {
                            pat.Invoke();
                            return true;
                        }
                        catch (COMException ex)
                        {
                            // Если именно 0x80040200 — сразу физический клик
                            if ((uint)ex.HResult == 0x80040200u)
                                return ClickElementPhysically(element);

                            // Иначе мягкий фолбэк: пробуем Legacy, потом мышь
                        }
                    }

                    // LegacyIAccessible как запасной вариант
                    var legacy = element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
                    if (legacy != null)
                    {
                        try
                        {
                            legacy.DoDefaultAction();
                            return true;
                        }
                        catch (COMException ex)
                        {
                            if ((uint)ex.HResult == 0x80040200u)
                                return ClickElementPhysically(element);
                        }
                    }

                    // Если нет паттернов или не сработало — физический клик
                    return ClickElementPhysically(element);
                }
                catch (COMException ex)
                {
                    // На транзиентных ошибках даём шанс ещё раз
                    if ((uint)ex.HResult == 0x80040200u || (uint)ex.HResult == 0x802A0001u)
                    {
                        Thread.Sleep(retryDelayMs);
                        continue;
                    }
                    // Иные ошибки — тоже пробуем мышью
                    if (ClickElementPhysically(element)) return true;
                    Thread.Sleep(retryDelayMs);
                }
                catch
                {
                    if (ClickElementPhysically(element)) return true;
                    Thread.Sleep(retryDelayMs);
                }
            }

            return false;
        }

        // --- ФИЗИЧЕСКИЙ КЛИК МЫШЬЮ ПО ЦЕНТРУ ЭЛЕМЕНТА ---
        static bool ClickElementPhysically(IUIAutomationElement el)
        {
            try
            {
                // Поднять окно на передний план, если возможно
                var hwndObj = el.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
                int hwnd = (hwndObj is int) ? (int)hwndObj : 0;
                if (hwnd != 0)
                {
                    ShowWindow(new IntPtr(hwnd), 5 /*SW_SHOW*/);
                    SetForegroundWindow(new IntPtr(hwnd));
                }

                // Координаты центра прямоугольника элемента
                var rectObj = el.GetCurrentPropertyValue(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
                if (!(rectObj is tagRECT)) return false;
                var r = (tagRECT)rectObj;
                if (r.right <= r.left || r.bottom <= r.top) return false;

                int x = r.left + (r.right - r.left) / 2;
                int y = r.top + (r.bottom - r.top) / 2;

                // Переместить курсор и клик ЛКМ
                SetCursorPos(x, y);
                Thread.Sleep(30);
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(10);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                return true;
            }
            catch
            {
                return false;
            }
        }
        static bool WaitWindowClosedByName(IUIAutomationElement parent, string name, int timeoutMs, int pollMs)
        {
            var automation = new CUIAutomation();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var cond = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name);
                    var el = parent.FindFirst(TreeScope.TreeScope_Descendants, cond);
                    if (el == null) return true; // закрылось
                }
                catch { /* игнор */ }
                Thread.Sleep(pollMs);
            }
            return false;
        }

        // Принудительное закрытие окна по Name (WindowPattern.Close -> WM_CLOSE -> Alt+F4)
        static void ForceCloseWindowByName(IUIAutomationElement parent, string name)
        {
            var automation = new CUIAutomation();
            try
            {
                var cond = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name);
                var el = parent.FindFirst(TreeScope.TreeScope_Descendants, cond);
                if (el == null) return;

                // 1) WindowPattern.Close
                try
                {
                    var wp = el.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId) as IUIAutomationWindowPattern;
                    if (wp != null)
                    {
                        wp.Close();
                        return;
                    }
                }
                catch { }

                // 2) WM_CLOSE по HWND
                var hwndObj = el.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
                int hwnd = (hwndObj is int) ? (int)hwndObj : 0;
                if (hwnd != 0)
                {
                    SendMessage(new IntPtr(hwnd), WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    return;
                }

                // 3) Alt+F4
                TryBringToFront(el);
                PressAltF4();
            }
            catch { /* не критично */ }
        }

        private const int SW_SHOWMAXIMIZED = 3;
        private static void TryMaximizeByWinApi(IUIAutomationElement element)
        {
            // Получаем HWND из свойства NativeWindowHandle
            object hwndObj = null;
            try
            {
                hwndObj = element.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
            }
            catch { }

            int hwndInt = 0;
            if (hwndObj is int) hwndInt = (int)hwndObj;

            if (hwndInt != 0)
            {
                IntPtr hWnd = (IntPtr)hwndInt;
                try { SetForegroundWindow(hWnd); } catch { }
                ShowWindow(hWnd, SW_SHOWMAXIMIZED);
            }
            else
            {
                // Если вдруг хэндл получить не удалось — ничего страшного, просто выходим
                // (окно уже найдено и в фокусе, этим можно ограничиться)
            }
        }

        // Поднять окно на передний план (если есть hwnd)
        static void TryBringToFront(IUIAutomationElement el)
        {
            try
            {
                var hwndObj = el.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
                int hwnd = (hwndObj is int) ? (int)hwndObj : 0;
                if (hwnd != 0)
                {
                    ShowWindow(new IntPtr(hwnd), 5); // SW_SHOW
                    SetForegroundWindow(new IntPtr(hwnd));
                }
            }
            catch { }
        }
        static void TryRenameWithRetries(string src, string dst, int attempts, int delayMs)
        {
            for (int i = 1; i <= attempts; i++)
            {
                try
                {
                    File.Move(src, dst);
                    Log(LogLevel.Info, $"[*] Помечен как обработанный: {Path.GetFileName(dst)}");
                    return;
                }
                catch (IOException ioEx)
                {
                    if (i == attempts)
                    {
                        Log(LogLevel.Warning, $"[!] Не удалось переименовать с добавлением ' ОК': {ioEx.Message}");
                        return;
                    }
                    System.Threading.Thread.Sleep(delayMs);
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Warning, $"[!] Переименование прервано: {ex.Message}");
                    return;
                }
            }
        }

        static void SafeCloseWindow(IUIAutomationElement window, string tag)
        {
            if (window == null) return;

            try
            {
                var patObj = window.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId);
                var winPat = patObj as IUIAutomationWindowPattern;
                if (winPat != null)
                {
                    winPat.Close();
                    Log(LogLevel.Info, "[" + tag + "] Закрываем окно через WindowPattern.Close().");
                }
            }
            catch { }

            try
            {
                try { window.SetFocus(); } catch { }
                System.Windows.Forms.SendKeys.SendWait("%{F4}");
                Log(LogLevel.Debug, "[" + tag + "] Отправлен Alt+F4.");
            }
            catch { }

            try
            {
                object hwndObj = window.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
                int hwndInt = (hwndObj is int) ? (int)hwndObj : 0;
                if (hwndInt != 0)
                {
                    var hWnd = (IntPtr)hwndInt;
                    if (Win32.IsWindow(hWnd))
                    {
                        IntPtr res;
                        Win32.SendMessageTimeout(hWnd, Win32.WM_SYSCOMMAND, new IntPtr(Win32.SC_CLOSE), IntPtr.Zero,
                            Win32.SMTO_ABORTIFHUNG, 2000, out res);
                        Win32.SendMessageTimeout(hWnd, Win32.WM_CLOSE, IntPtr.Zero, IntPtr.Zero,
                            Win32.SMTO_ABORTIFHUNG, 2000, out res);
                        Log(LogLevel.Debug, "[" + tag + "] Отправлены SC_CLOSE/WM_CLOSE.");
                    }
                }
            }
            catch { }
        }

        static void WaitWindowGoneByHandle(IUIAutomationElement window, int timeoutMs)
        {
            if (window == null) return;
            IntPtr hWnd = IntPtr.Zero;
            try
            {
                object hwndObj = window.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
                int hwndInt = (hwndObj is int) ? (int)hwndObj : 0;
                if (hwndInt != 0) hWnd = (IntPtr)hwndInt;
            }
            catch { }

            if (hWnd == IntPtr.Zero) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (!Win32.IsWindow(hWnd)) break;
                System.Threading.Thread.Sleep(100);
            }
            sw.Stop();
        }


        // Нажать Enter глобально
        static void PressEnter()
        {
            INPUT[] inputs = new INPUT[2];
            inputs[0].type = 1; // KEYBOARD
            inputs[0].U.ki = new KEYBDINPUT { wVk = VK_RETURN };
            inputs[1].type = 1;
            inputs[1].U.ki = new KEYBDINPUT { wVk = VK_RETURN, dwFlags = KEYEVENTF_KEYUP };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        // Нажать Alt+F4 глобально
        static void PressAltF4()
        {
            INPUT[] inputs = new INPUT[4];
            inputs[0].type = 1; inputs[0].U.ki = new KEYBDINPUT { wVk = VK_MENU };                       // Alt down
            inputs[1].type = 1; inputs[1].U.ki = new KEYBDINPUT { wVk = VK_F4 };                         // F4 down
            inputs[2].type = 1; inputs[2].U.ki = new KEYBDINPUT { wVk = VK_F4, dwFlags = KEYEVENTF_KEYUP }; // F4 up
            inputs[3].type = 1; inputs[3].U.ki = new KEYBDINPUT { wVk = VK_MENU, dwFlags = KEYEVENTF_KEYUP }; // Alt up
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        // P/Invoke и структуры для SendInput/WM_CLOSE/Foreground
        const int WM_CLOSE = 0x0010;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const byte VK_RETURN = 0x0D;
        const byte VK_MENU = 0x12; // Alt
        const byte VK_F4 = 0x73;

        [DllImport("user32.dll")] static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx, dy;
            public uint mouseData, dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg, wParamL, wParamH;
        }

        internal static class Win32
        {
            public const uint WM_CLOSE = 0x0010;
            public const uint WM_SYSCOMMAND = 0x0112;
            public const int SC_CLOSE = 0xF060;
            public const uint SMTO_ABORTIFHUNG = 0x0002;

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SendMessageTimeout(
                IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
                uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

            [DllImport("user32.dll")]
            public static extern bool IsWindow(IntPtr hWnd);
        }
    }
}

using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace My_Console
{
    class Program
    {
        private const string EmailXPath_Address = "B1";
        private const string SubmitBtnXPath_Address = "B2";
        private const string LinkForm_Address = "B3";
        private static string _emailXPath;
        private static string _submitBtnXPath;
        private static string _linkForm;
        private static List<string> _xpaths;

        static async Task Main(string[] args)
        {
            var data = GetData("Data.xlsx");
            var currentUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            var chromeUserDataPath = $@"C:\Users\{currentUser}\AppData\Local\Google\Chrome\User Data\";

            var chromeUsers = Directory.GetDirectories(chromeUserDataPath)
                .Select(fullName => fullName.Split('\\').Last())
                .Where(folderName => Regex.IsMatch(folderName, "^Profile [1-9]{1,}$"));


            foreach (var chromeUser in chromeUsers)
            {
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument($"user-data-dir={chromeUserDataPath}");
                chromeOptions.AddArgument($"profile-directory={chromeUser}");
                var chromeDriver = new ChromeDriver(chromeOptions);
                chromeDriver.Url = _linkForm;
                chromeDriver.Navigate();
                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(2));

                    //var clearButton = chromeDriver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div/div[3]/div[1]/div[2]/div"));
                    //if (clearButton != null)
                    //{
                    //    clearButton.Click();
                    //    await Task.Delay(TimeSpan.FromSeconds(2));
                    //    var confirmButton = chromeDriver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/div[3]/div[2]"));
                    //    confirmButton.Click();
                    //    await Task.Delay(TimeSpan.FromSeconds(2));
                    //}

                    var emailAddress = chromeDriver.FindElement(By.XPath(_emailXPath)).Text;

                    var currentData = data.SingleOrDefault(d => d.Last() == chromeUser);
                    if (currentData == null)
                    {
                        continue;
                    }

                    for (var xpathIndex = 0; xpathIndex < _xpaths.Count - 1; xpathIndex++)
                    {
                        try
                        {
                            var xpath = _xpaths.ElementAt(xpathIndex);
                            var fieldValue = currentData.ElementAt(xpathIndex);
                            if (fieldValue.ToLower().Equals("@email"))
                            {
                                fieldValue = emailAddress;
                            }

                            chromeDriver.FindElement(By.XPath(xpath)).SendKeys(fieldValue);
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    var btnSubmit = chromeDriver.FindElement(By.XPath(_submitBtnXPath));
                    btnSubmit.Click();
                    await Task.Delay(TimeSpan.FromSeconds(3));
                }
                catch (Exception)
                {
                }
                finally
                {
                    await Task.Delay(TimeSpan.FromSeconds(3));
                    chromeDriver.Close();
                    chromeDriver.Quit();
                }
            }

            Console.Read();
        }

        static IEnumerable<IEnumerable<string>> GetData(string dataFilePath)
        {
            var excelPackage = new ExcelPackage(File.OpenRead(dataFilePath));
            var worksheet = excelPackage.Workbook.Worksheets.First();
            var maxCol = worksheet.Dimension.Columns;
            var maxRow = worksheet.Dimension.Rows;

            _emailXPath = worksheet.Cells[EmailXPath_Address].Value.ToString();
            _submitBtnXPath = worksheet.Cells[SubmitBtnXPath_Address].Value.ToString();
            _linkForm = worksheet.Cells[LinkForm_Address].Value.ToString();

            var data = new List<List<string>>();
            for (var rowIndex = 4; rowIndex <= maxRow; rowIndex++)
            {
                var colsData = new List<string>();
                for (var colIndex = 1; colIndex <= maxCol; colIndex++)
                {
                    (rowIndex == 4 ? _xpaths ??= new List<string>() : colsData).Add(worksheet.Cells[rowIndex, colIndex].Value?.ToString());
                }

                if (rowIndex != 4)
                {
                    data.Add(colsData);
                }
            }

            return data;
        }
    }
}

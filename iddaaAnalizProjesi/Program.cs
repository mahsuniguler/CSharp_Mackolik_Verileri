using HtmlAgilityPack;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace iddaaAnalizProjesi
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ChromeOptions options = new ChromeOptions();

            IWebDriver driver = new ChromeDriver(options);

            string url = "https://arsiv.mackolik.com/Genis-Iddaa-Programi";
            driver.Navigate().GoToUrl(url);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("justNotPlayed")));
            driver.FindElement(By.Id("justNotPlayed")).Click();

            System.Threading.Thread.Sleep(1000);

            // Element bulundu, bekleme koşulu kullanabilirsiniz.
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("dayId")));
            SelectElement listbox = new SelectElement(driver.FindElement(By.Id("dayId")));
            listbox.SelectByText("Hepsi");


            string table_class = "iddaa-oyna-table";
            //wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName(table_class)));

            IWebElement table = driver.FindElement(By.ClassName(table_class));
            table = driver.FindElement(By.ClassName(table_class));

            string tableHtml = table.GetAttribute("outerHTML");

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(tableHtml);

            var rows = htmlDoc.DocumentNode.SelectNodes("//tr");
            List<List<string>> tabloVerileri = new List<List<string>>();
            List<string> resimler = new List<string>(); 

            foreach (var row in rows)
            {
                var hücreler = row.SelectNodes("td");
                if (hücreler != null)
                {
                    List<string> satırVerileri = new List<string>();
                    foreach (var hücre in hücreler)
                    {
                        satırVerileri.Add(hücre.InnerText);

                    }
                    tabloVerileri.Add(satırVerileri);
                }
            }


            int dosyaNumarasi = 0;
            string excelDosyaAdi = $"C:/Users/Mahsuni/Desktop/Analiz/veriler{dosyaNumarasi}.xlsx";


            // Dosya adının varlığını ve kullanılabilirliğini kontrol edin.
            while (File.Exists(excelDosyaAdi))
            {
                dosyaNumarasi++;
                excelDosyaAdi = $"C:/Users/Mahsuni/Desktop/Analiz/veriler{dosyaNumarasi}.xlsx";

            }
            
            DateTime bugununTarihi = DateTime.Now;
            string bugununTarihistr = bugununTarihi.ToString("dd.MM.yyyy");
            OfficeOpenXml.ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelDosyaYolu = excelDosyaAdi;

            using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelDosyaAdi)))
            {
                var worksheet = package.Workbook.Worksheets.Add(bugununTarihistr);

                for (int satırIndex = 0; satırIndex < tabloVerileri.Count; satırIndex++)
                {
                    int alSutunIndex = 37;

                    if (tabloVerileri[satırIndex].Count > alSutunIndex)
                    {
                        tabloVerileri[satırIndex].RemoveAt(alSutunIndex);
                    }
                    int satirsayisi=tabloVerileri[satırIndex].Count;
                    for (int sütunIndex = 0; sütunIndex < satirsayisi; sütunIndex++)
                    {
                        if (tabloVerileri[satırIndex][sütunIndex] == "" || tabloVerileri[satırIndex][sütunIndex]==" ") tabloVerileri[satırIndex][sütunIndex] = "";
                        worksheet.Cells[satırIndex + 1, sütunIndex + 1].Value = tabloVerileri[satırIndex][sütunIndex].Replace("   &nbsp;", "").Trim();
                        worksheet.Cells[satırIndex + 1, sütunIndex + 1].Value = tabloVerileri[satırIndex][sütunIndex].Replace("  &nbsp;", "").Trim();
                        worksheet.Cells[satırIndex + 1, sütunIndex + 1].Value = tabloVerileri[satırIndex][sütunIndex].Replace(" &nbsp;", "").Trim();
                        worksheet.Cells[satırIndex + 1, sütunIndex + 1].Value = tabloVerileri[satırIndex][sütunIndex].Replace("&nbsp;", "").Trim();
                    }
                }

                foreach (var row in rows)
                {

                    var img = row.SelectSingleNode("td[5]/img");
                    if (img != null)
                    {
                        string src = img.GetAttributeValue("src", "");
                        if (src.EndsWith("1.gif") || src.EndsWith("2.gif") || src.EndsWith("3.gif"))
                            resimler.Add(src.Substring(src.Length - 5,1));
                        else
                            resimler.Add("-");
                    }
                    else
                    {
                        resimler.Add("");
                    }
                }
                for (int i = 4; i < tabloVerileri.Count; i++)
                {
                    worksheet.Cells[i+ 1,5].Value = resimler[i];
                }


                package.Save();
                Console.WriteLine($"{excelDosyaAdi} kaydedildi.");
            }

            driver.Quit();
        }
    }
}

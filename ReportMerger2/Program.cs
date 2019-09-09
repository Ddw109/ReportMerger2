using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ReportMerger2.Model;
using Microsoft.Extensions.Configuration;
using Atp.Pdf;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace ReportMerger2
{
    class Program
    {
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");
            var Configuration = builder.Build();
            var appConfig = Configuration.GetSection("App").Get<Config>();

            if (DateTime.Today.DayOfWeek != DayOfWeek.Saturday && DateTime.Today.DayOfWeek != DayOfWeek.Sunday)
            {
                var currentTime = ((DateTime.Today.DayOfWeek == DayOfWeek.Monday) ? DateTime.Now.AddDays(-3).ToString(appConfig.dateFormat) : DateTime.Now.AddDays(-1).ToString(appConfig.dateFormat));

                //Sales_MonthlyToBudget
                ExcelToPdf(DateTime.Now.Year.ToString(), appConfig.ExcelQueriesLocation, appConfig.Sales_MonthlySalesToBudget_xls, appConfig.Sales_MonthlySalesToBudget_pdf);

                // Log Report
                Merger(appConfig.Log, appConfig.LogDaily, appConfig.sourcePath, appConfig.dropPath, currentTime, appConfig.LogisticsDailyReports);
                // Trans Report
                ExcelToPdf(appConfig.TransDailySheet, appConfig.sourcePath, appConfig.TransDailyExcel, appConfig.TransDaily.ElementAt(0));
                Merger(appConfig.Trans, appConfig.TransDaily, appConfig.sourcePath, appConfig.dropPath, currentTime, appConfig.TransDailyReports);
            }
        }

        private static void Merger(string reportType, List<string> reports, string sourcePath, string dropPath, string currentTime, string reportName)
        {
            var combineFile = false;
            PdfDocument Pdf = new PdfDocument();
            Console.WriteLine(String.Format("Starting {0} Report Merging...", reportType));
            foreach (string file in reports)
            {
                if (File.Exists(sourcePath + file))
                {
                    Console.WriteLine("Appending : " + file);
                    combineFile = true;
                    Pdf.Append(new PdfImportedDocument(sourcePath + file));
                }
            }
            if (combineFile)
            {
                combineFile = false;
                Console.WriteLine("Saving : " + String.Format(reportName, currentTime));
                Pdf.Save(dropPath + String.Format(reportName, currentTime));
                Console.WriteLine("Done");
            }
            Pdf.Close();
        }

        private static void ExcelToPdf(string sheetName,string location , string inFileName, string outFileName)
        {
            Application app = new Application();
            Console.WriteLine("Converting " + inFileName + " to " + outFileName);
            Thread.Sleep(1000 * 20);
            Workbook wkb = app.Workbooks.Open(location + inFileName, ReadOnly: true);
            wkb.Worksheets[sheetName].ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, location + outFileName);
            Console.WriteLine("Done");            
        }
    }
}

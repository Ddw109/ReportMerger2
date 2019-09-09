using System.Collections.Generic;

namespace ReportMerger2.Model
{
    class Config
    {
        public string sourcePath { get; set; }
        public string dropPath { get; set; }
        public string Log { get; set; }
        public List<string> LogDaily { get; set; }
        public string LogisticsDailyReports { get; set; }
        public string Trans { get; set; }
        public string TransDailyExcel { get; set; }
        public string TransDailySheet { get; set; }
        public List<string> TransDaily { get; set; }
        public string TransDailyReports { get; set; }
        public string dateFormat { get; set; }
        public string ExcelQueriesLocation { get; set; }
        public string Sales_MonthlySalesToBudget_xls { get; set; }
        public string Sales_MonthlySalesToBudget_pdf { get; set; }
    }
}

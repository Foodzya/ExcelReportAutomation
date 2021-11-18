using Spire.Xls;
using System;
using System.IO;
using System.Linq;

namespace ReportAutomatizator
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelManipulation.ConvertCsvToXlsx();

            ExcelManipulation.CopyPasteSheetToNewWorkbook("Total_Sol.xlsx");

            ExcelManipulation.CalculateValuesFromDataRegistrationSheet("No-affiliate", "ОКТЯБРЬ");
        } 
    }

    static class ExcelManipulation 
    {
        public static void CopyPasteSheetToNewWorkbook(string fileName)
        {
            // Load source Excel file.
            Workbook sourceBook = new Workbook();
            sourceBook.LoadFromFile($"C:\\Users\\a.borisyonok\\Downloads\\Excel\\{fileName}");

            Worksheet sourceSheet = sourceBook.Worksheets[0];
            
            CellRange sourceRange = sourceSheet.Range[sourceSheet.FirstRow, sourceSheet.FirstColumn, sourceSheet.LastRow, sourceSheet.LastColumn];

            Workbook destBook = new Workbook();
            destBook.LoadFromFile(@"C:\Users\a.borisyonok\Downloads\Excel\Продуктовый отчёт.xlsx");

            Worksheet destSheet = (Worksheet)destBook.Worksheets.First(w => w.Name == "Техническая (дата регистрации)");

            // Deletes last data from the destination source.
            destSheet.Range[destSheet.FirstRow, destSheet.FirstColumn, destSheet.LastRow, 84].Clear(ExcelClearOptions.ClearContent);

            CellRange destRange = destSheet.Range[sourceSheet.FirstRow, sourceSheet.FirstColumn, sourceSheet.LastRow, sourceSheet.LastColumn];

            sourceSheet.Copy(sourceRange, destRange);

            destBook.SaveToFile(@"C:\Users\a.borisyonok\Downloads\Excel\Продуктовый отчёт.xlsx");   
        }

        public static void CalculateValuesFromDataRegistrationSheet(string value, string reportingMonth)
        {
            // int numberOfGroupColumn = 18;

            Workbook summaryReport = new Workbook();
            summaryReport.LoadFromFile(@"C:\Users\a.borisyonok\Downloads\Excel\Продуктовый отчёт.xlsx");

            Worksheet sourceSheet = (Worksheet)summaryReport.Worksheets.First(p => p.Name == "Техническая (дата регистрации)");

            // Detect number of rows in email [2] column (number of customers).
            int lastNotBlankRow = sourceSheet.Columns[2].Rows.Where(cs => !cs.IsBlank).Count() - 1;

            // Finds the number of a column with the name "Группы".
            int numberOfGroupColumn = sourceSheet.Range[sourceSheet.FirstRow, sourceSheet.FirstColumn, sourceSheet.FirstRow, sourceSheet.LastColumn].Columns.First(c => c.Value == "Группы").Column;

            // Getting the range of 18th column ("Группа").
            CellRange range = sourceSheet.Range[sourceSheet.FirstRow, numberOfGroupColumn, lastNotBlankRow, numberOfGroupColumn];

            // Setting value for blank cells in column "Группа".
            range.Where(r => r.IsBlank).Select(r => r.Value = value).ToList();    

            summaryReport.CalculateAllValue();      

            Worksheet reportMonthSheet = (Worksheet)summaryReport.Worksheets.First(w => w.Name == $"{reportingMonth}");

            var dateOfRegistration = Convert.ToDateTime(sourceSheet.Rows[2].CellList[13].Value);

            CellRange dateRange = reportMonthSheet.Range[2, reportMonthSheet.FirstColumn, 2, reportMonthSheet.LastColumn];

            int destColIndex = 0;

            // This loop checks for registration date of a daily report
            // then finds the column with such date in main report sheet
            // for the future copy-paste as value.
            foreach(var cellValue in dateRange)
            {
                DateTime cellDate = default;              

                if(!cellValue.IsBlank && cellDate.Day == dateOfRegistration.Day && cellDate.Month == dateOfRegistration.Month && cellDate.Year == dateOfRegistration.Year)
                {
                    cellDate = Convert.ToDateTime(cellValue.Value);

                    // Index of destination column (which needs to be copy-paste as value).
                    destColIndex = cellValue.Column;
                    break;
                }
            }

            CellRange columnToBePastAsValue = reportMonthSheet.Range[reportMonthSheet.FirstRow, destColIndex, reportMonthSheet.LastRow, destColIndex];

            reportMonthSheet.CopyColumn(columnToBePastAsValue, reportMonthSheet, destColIndex, CopyRangeOptions.OnlyCopyFormulaValue);     

            summaryReport.Save();
        }

        public static void ConvertCsvToXlsx()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile($"C:\\Users\\a.borisyonok\\Downloads\\Excel\\Total_Sol.csv", ";", 1, 1, ExcelVersion.Version2016, System.Text.Encoding.UTF8);

            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Csv to Xlsx";

            workbook.SaveToFile(@"C:\Users\a.borisyonok\Downloads\Excel\Total_Sol.xlsx", ExcelVersion.Version2016);

            File.Delete($"C:\\Users\\a.borisyonok\\Downloads\\Excel\\Total_Sol.csv");
        }
    }
}

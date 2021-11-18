using Spire.Xls;
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

            Worksheet destSheet = destBook.Worksheets[3];

            // Deletes last data from the destination source.
            destSheet.Range[destSheet.FirstRow, destSheet.FirstColumn, destSheet.LastRow, 84].Clear(ExcelClearOptions.ClearContent);

            CellRange destRange = destSheet.Range[sourceSheet.FirstRow, sourceSheet.FirstColumn, sourceSheet.LastRow, sourceSheet.LastColumn];

            destBook.Save();

            sourceSheet.Copy(sourceRange, destRange);

            destBook.SaveToFile(@"C:\Users\a.borisyonok\Downloads\Excel\Продуктовый отчёт.xlsx");   
        }

        public static void CalculateValuesFromDataRegistrationSheet(string value, string reportingMonth)
        {
            int numberOfGroupColumn = 18;

            Workbook sourceBook = new Workbook();
            sourceBook.LoadFromFile(@"C:\Users\a.borisyonok\Downloads\Excel\Продуктовый отчёт.xlsx");

            Worksheet sourceSheet = (Worksheet)sourceBook.Worksheets.First(p => p.Name == "Техническая (дата регистрации)");

            // Detect number of rows in email [2] column (number of customers).
            int lastNotBlankRow = sourceSheet.Columns[2].Rows.Where(cs => !cs.IsBlank).Count() - 1;

            // Getting the range of 18th column ("Группа").
            CellRange range = sourceSheet.Range[sourceSheet.FirstRow, numberOfGroupColumn, lastNotBlankRow, numberOfGroupColumn];

            // Setting value for blank cells in column "Группа".
            range.Where(r => r.IsBlank).Select(r => r.Value = value).ToList();    

            sourceBook.CalculateAllValue();      

            Worksheet reportMonthSheet = (Worksheet)sourceBook.Worksheets.First(w => w.Name == $"{reportingMonth}");

            // Index of destination column (which needs to be copy-paste as value).
            int destColIndex = 12;

            CellRange columnToBePastAsValue = reportMonthSheet.Range[reportMonthSheet.FirstRow, destColIndex, reportMonthSheet.LastRow, destColIndex];

            reportMonthSheet.CopyColumn(columnToBePastAsValue, reportMonthSheet, destColIndex, CopyRangeOptions.OnlyCopyFormulaValue);     

            sourceBook.Save();
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

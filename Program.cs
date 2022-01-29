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

            ExcelManipulation.CopyPasteSheetToNewWorkbook("Total_Sol.xlsx", ExcelManipulation.GetMainReportFile("Продуктовый отчёт.xlsx"));

            ExcelManipulation.CalculateValuesFromDataRegistrationSheet("No-affiliate", "ОКТЯБРЬ", ExcelManipulation.GetMainReportFile("Продуктовый отчёт.xlsx"));
        } 
    }

    static class ExcelManipulation 
    {
        public static Workbook GetMainReportFile(string dailyReportFile)
        {
            Workbook sourceBook = new Workbook();

            sourceBook.LoadFromFile($"C:\\Users\\a.borisyonok\\Downloads\\Excel\\{dailyReportFile}");

            return sourceBook;
        }

        public static void CopyPasteSheetToNewWorkbook(string dailyReportFile, Workbook mainReportBook)
        {
            // Exact number of columns which are involved in replacement process.
            int numberOfInvolvedColumns = 84;

            // Load source Excel file.
            Workbook sourceBook = new Workbook();
            sourceBook.LoadFromFile($"C:\\Users\\a.borisyonok\\Downloads\\Excel\\{dailyReportFile}");

            Worksheet dailyReportSheet = sourceBook.Worksheets[0];
            
            CellRange sourceRange = dailyReportSheet.Range[dailyReportSheet.FirstRow, dailyReportSheet.FirstColumn, dailyReportSheet.LastRow, dailyReportSheet.LastColumn];

            Worksheet destSheet = (Worksheet)mainReportBook.Worksheets.First(w => w.Name == "Техническая (дата регистрации)");

            // Deletes last data from the destination source.
            // It needs when previous report had more lines than the current one.
            destSheet.Range[destSheet.FirstRow, destSheet.FirstColumn, destSheet.LastRow, numberOfInvolvedColumns].Clear(ExcelClearOptions.ClearContent);

            CellRange destRange = destSheet.Range[dailyReportSheet.FirstRow, dailyReportSheet.FirstColumn, dailyReportSheet.LastRow, dailyReportSheet.LastColumn];

            dailyReportSheet.Copy(sourceRange, destRange);

            mainReportBook.Save();  
        }

        // public static void CalculateValuesFromDataActionSheet()

        public static void CalculateValuesFromDataRegistrationSheet(string valueForBlankFields, string reportMonthName, Workbook mainReportBook)
        {
            Worksheet sourceSheet = (Worksheet)mainReportBook.Worksheets.First(p => p.Name == "Техническая (дата регистрации)");

            // Detect number of rows in email [2] column (number of customers).
            int lastNotBlankRow = sourceSheet.Columns[2].Rows.Where(cs => !cs.IsBlank).Count() - 1;

            // Finds the number of a column with the name "Группы".
            int numberOfGroupColumn = sourceSheet.Range[sourceSheet.FirstRow, sourceSheet.FirstColumn, sourceSheet.FirstRow, sourceSheet.LastColumn].Columns.First(c => c.Value == "Группы").Column;

            // Getting the range of 18th column ("Группа").
            CellRange range = sourceSheet.Range[sourceSheet.FirstRow, numberOfGroupColumn, lastNotBlankRow, numberOfGroupColumn];

            // Setting value for blank cells in column "Группа".
            range.Where(r => r.IsBlank).Select(r => r.Value = valueForBlankFields).ToList();    

            mainReportBook.CalculateAllValue();      

            Worksheet reportMonthSheet = (Worksheet)mainReportBook.Worksheets.First(w => w.Name == $"{reportMonthName}");

            var dateOfRegistration = Convert.ToDateTime(sourceSheet.Rows[2].CellList[13].Value);

            CellRange dateRange = reportMonthSheet.Range[2, reportMonthSheet.FirstColumn, 2, reportMonthSheet.LastColumn];

            int destColIndex = 0;

            // This loop checks for registration date of a daily report
            // then finds the column with such date in main report sheet
            // to leave only values in cells in further.
            foreach(var cellValue in dateRange)
            {
                DateTime cellDate = default;   

                if(!cellValue.IsBlank) 
                {
                    cellDate = Convert.ToDateTime(cellValue.Value);

                    if(cellDate.Day == dateOfRegistration.Day && cellDate.Month == dateOfRegistration.Month && cellDate.Year == dateOfRegistration.Year)
                    {
                        // Index of destination column (which needs to be copy-paste as value).
                        destColIndex = cellValue.Column;
                        break;
                    }
                }           
            }

            CellRange currentReportDayColumn = reportMonthSheet.Range[reportMonthSheet.FirstRow, destColIndex, reportMonthSheet.LastRow, destColIndex];

            foreach(CellRange cell in currentReportDayColumn)
            {
                if(cell.HasFormula)
                {
                    var cellValue = cell.FormulaValue;

                    cell.Clear(ExcelClearOptions.ClearContent);

                    cell.Value2 = cellValue;
                }
            }
            mainReportBook.Save();
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

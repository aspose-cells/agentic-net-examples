using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ClearPrintTitlesDemo
    {
        public static void Run()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            PageSetup pageSetup = worksheet.PageSetup;

            // Set print title rows and columns (these will repeat on each printed page)
            pageSetup.PrintTitleRows = "$1:$2";      // repeat rows 1‑2
            pageSetup.PrintTitleColumns = "$A:$B";   // repeat columns A‑B

            Console.WriteLine("Print titles set:");
            Console.WriteLine($"Rows: {pageSetup.PrintTitleRows}");
            Console.WriteLine($"Columns: {pageSetup.PrintTitleColumns}");

            // Clear the print title rows and columns by assigning an empty string
            pageSetup.PrintTitleRows = string.Empty;
            pageSetup.PrintTitleColumns = string.Empty;

            Console.WriteLine("\nPrint titles cleared:");
            Console.WriteLine($"Rows: {(string.IsNullOrEmpty(pageSetup.PrintTitleRows) ? "Cleared" : pageSetup.PrintTitleRows)}");
            Console.WriteLine($"Columns: {(string.IsNullOrEmpty(pageSetup.PrintTitleColumns) ? "Cleared" : pageSetup.PrintTitleColumns)}");

            // Save the workbook to verify the changes
            workbook.Save("ClearPrintTitlesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ClearPrintTitlesDemo.Run();
        }
    }
}
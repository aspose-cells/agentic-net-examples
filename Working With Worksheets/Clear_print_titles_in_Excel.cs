using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ClearPrintTitlesDemo
    {
        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set some print titles to demonstrate that they exist
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.PrintTitleRows = "$1:$2";      // repeat rows 1-2 on each printed page
            pageSetup.PrintTitleColumns = "$A:$B";   // repeat columns A-B on each printed page

            Console.WriteLine("Before clearing:");
            Console.WriteLine($"PrintTitleRows = {pageSetup.PrintTitleRows}");
            Console.WriteLine($"PrintTitleColumns = {pageSetup.PrintTitleColumns}");

            // Clear the print titles by assigning empty strings
            pageSetup.PrintTitleRows = string.Empty;
            pageSetup.PrintTitleColumns = string.Empty;

            Console.WriteLine("\nAfter clearing:");
            Console.WriteLine($"PrintTitleRows = {(string.IsNullOrEmpty(pageSetup.PrintTitleRows) ? "Cleared" : pageSetup.PrintTitleRows)}");
            Console.WriteLine($"PrintTitleColumns = {(string.IsNullOrEmpty(pageSetup.PrintTitleColumns) ? "Cleared" : pageSetup.PrintTitleColumns)}");

            // Save the workbook
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
using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ClearPrintTitlesDemo
    {
        public static void Run()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            for (int row = 0; row < 20; row++)
                for (int col = 0; col < 5; col++)
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");

            // Set print title rows and columns (e.g., first row and first column)
            worksheet.PageSetup.PrintTitleRows = "$1:$1";
            worksheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Clear the rows and columns that are repeated on each printed page
            worksheet.PageSetup.PrintTitleRows = string.Empty;
            worksheet.PageSetup.PrintTitleColumns = string.Empty;

            // Optional: display the cleared values for verification
            Console.WriteLine($"PrintTitleRows after clear: '{worksheet.PageSetup.PrintTitleRows}'");
            Console.WriteLine($"PrintTitleColumns after clear: '{worksheet.PageSetup.PrintTitleColumns}'");

            // Save the modified workbook
            workbook.Save("OutputClearedPrintTitles.xlsx");
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
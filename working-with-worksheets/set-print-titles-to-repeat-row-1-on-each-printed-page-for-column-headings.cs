using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetPrintTitleRowsDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a header row (row 1) and some sample data
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            for (int i = 2; i <= 50; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data A{i - 1}");
                worksheet.Cells[$"B{i}"].PutValue($"Data B{i - 1}");
            }

            // Repeat the first row on each printed page
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Define the print area (optional, but clarifies the range to print)
            worksheet.PageSetup.PrintArea = "A1:B50";

            // Save the workbook
            workbook.Save("PrintTitleRowsDemo.xlsx");
        }
    }
}
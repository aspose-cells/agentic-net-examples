using System;
using Aspose.Cells;

namespace AsposeCellsPivotPreparation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the full path to the existing Excel file
            string filePath = @"C:\Data\SampleWorkbook.xlsx";

            // Load the workbook from the file path using the Workbook(string) constructor
            Workbook workbook = new Workbook(filePath);

            // At this point the workbook is loaded and ready for pivot table operations
            // Example: access the first worksheet (optional)
            Worksheet worksheet = workbook.Worksheets[0];
            Console.WriteLine($"Loaded workbook with {workbook.Worksheets.Count} worksheet(s).");
            Console.WriteLine($"First worksheet name: {worksheet.Name}");

            // No saving is performed here as the task only requires loading for further processing
        }
    }
}
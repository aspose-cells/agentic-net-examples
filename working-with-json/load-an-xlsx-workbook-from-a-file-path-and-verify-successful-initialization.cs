using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "example.xlsx";

            // Load the workbook using the constructor that accepts a file path
            // This follows the provided rule: Workbook(string)
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook was initialized successfully
            // A simple check is to ensure that at least one worksheet is present
            if (workbook != null && workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("Workbook loaded successfully.");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            }
            else
            {
                Console.WriteLine("Failed to load the workbook or it contains no worksheets.");
            }

            // Dispose the workbook when done (optional, as it implements IDisposable)
            workbook.Dispose();
        }
    }
}
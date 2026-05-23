using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThreadedCalcDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (empty workbook with a default worksheet)
                Workbook workbook = new Workbook();

                // Multi‑threaded calculation is enabled by default.
                // If you need to disable it, ensure the property exists in your Aspose.Cells version.
                // workbook.Settings.UseThreadedCalculation = false; // Uncomment if supported.

                // Add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].Formula = "=A1+A2";

                // Define output file path
                string outputPath = "SimpleWorkbook.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
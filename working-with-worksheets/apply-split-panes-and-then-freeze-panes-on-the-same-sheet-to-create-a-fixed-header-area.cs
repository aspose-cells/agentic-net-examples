using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SplitAndFreezeDemo
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data with a header row
                sheet.Cells["A1"].PutValue("Header");
                for (int i = 0; i < 100; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue($"Row {i + 1}");
                    sheet.Cells[i + 1, 1].PutValue(i * 10);
                }

                // Split the worksheet window (creates separate panes)
                sheet.Split();

                // Freeze the top pane so the header row stays visible while scrolling
                // Freeze at row index 1 (second row), column index 0, freezing 1 row and 0 columns
                sheet.FreezePanes(1, 0, 1, 0);

                // Define output file path
                string outputPath = "SplitAndFreezeDemo.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }
}
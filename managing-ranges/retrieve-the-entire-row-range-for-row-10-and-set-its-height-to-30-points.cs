using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class SetRowHeightDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the entire row range for row 10 (1‑based index)
                // "10:10" denotes the whole 10th row
                AsposeRange row10Range = worksheet.Cells.CreateRange("10:10");

                // Set the height of this row range to 30 points
                row10Range.RowHeight = 30;

                // Optionally, verify the height
                Console.WriteLine("Row 10 height set to: " + row10Range.RowHeight);

                // Define output file path
                string outputPath = "RowHeightDemo.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCutPasteDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet as the source sheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Add a second worksheet as the destination sheet
                Worksheet destinationSheet = workbook.Worksheets.Add("Destination");

                // Populate sample data in the source range B2:C4
                sourceSheet.Cells["B2"].PutValue("Item1");
                sourceSheet.Cells["C2"].PutValue(100);
                sourceSheet.Cells["B3"].PutValue("Item2");
                sourceSheet.Cells["C3"].PutValue(200);
                sourceSheet.Cells["B4"].PutValue("Item3");
                sourceSheet.Cells["C4"].PutValue(300);

                // Create the source range (B2:C4)
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("B2:C4");

                // Create the destination range (G5:H7) on the destination sheet
                AsposeRange destinationRange = destinationSheet.Cells.CreateRange("G5:H7");

                // Copy data (values, formulas, formatting, etc.) from source to destination
                destinationRange.CopyData(sourceRange);

                // Clear the original source range to complete the "cut" operation
                // B2 corresponds to row index 1, column index 1 (zero‑based)
                sourceSheet.Cells.ClearRange(1, 1, 3, 2); // 3 rows, 2 columns

                // Define output file path
                string outputPath = "CutPasteResult.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log exception details
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
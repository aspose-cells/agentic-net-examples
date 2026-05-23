using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyAndRemoveFormulas
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Define the source range to copy (e.g., A1:C10 on the first worksheet)
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C10");

                // Create a new (empty) destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Ensure the destination workbook has at least one worksheet
                Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

                // Define the destination range where the data will be copied
                AsposeRange destinationRange = destinationSheet.Cells.CreateRange("A1:C10");

                // Copy the source range data (including formulas) to the destination range
                destinationRange.CopyData(sourceRange);

                // Remove all formulas from the destination worksheet, leaving only static values
                destinationSheet.Cells.RemoveFormulas();

                // Save the resulting workbook
                destinationWorkbook.Save(outputPath);

                Console.WriteLine("Range copied and formulas removed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
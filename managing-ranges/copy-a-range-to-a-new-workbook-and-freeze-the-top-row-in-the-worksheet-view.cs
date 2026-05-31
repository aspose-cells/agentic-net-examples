using System;
using System.IO;
using Aspose.Cells;

class CopyRangeAndFreezeTopRow
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "output.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the source range to copy (e.g., A1:C10)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C10");

            // Create a new (empty) workbook for the destination
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Define the destination range where the data will be copied (starting at A1)
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:C10");

            // Copy the source range to the destination range (includes data, formatting, etc.)
            destRange.Copy(sourceRange);

            // Freeze the top row in the destination worksheet view (freeze at cell A2)
            destSheet.FreezePanes("A2", 1, 0);

            // Save the destination workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
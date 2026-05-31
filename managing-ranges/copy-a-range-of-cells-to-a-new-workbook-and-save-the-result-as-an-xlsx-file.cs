using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "copied_range.xlsx";

            // Verify that the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the source range to copy (e.g., A1:C5)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define a destination range with the same dimensions
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("A1:C5");

            // Copy the source range (including data, formulas, formatting, etc.) to the destination range
            destinationRange.Copy(sourceRange);

            // Save the destination workbook
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Range copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
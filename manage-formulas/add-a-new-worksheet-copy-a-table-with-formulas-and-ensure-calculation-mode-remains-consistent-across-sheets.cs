// Title: Copy a formula‑rich table to a newly added worksheet while preserving the workbook’s calculation mode with Aspose.Cells for .NET
// Common Searches: asp.net copy range with formulas to a new worksheet using Aspose.Cells | how to keep Excel calculation mode unchanged after adding a sheet in Aspose.Cells | duplicate a table that contains formulas without breaking references Aspose.Cells C# | preserve workbook calculation settings while copying data between worksheets Aspose.Cells | copy Excel range preserving formulas and calculation mode programmatically in C#
// Tags: copy range with formulas Aspose.Cells | add worksheet duplicate table C# | preserve workbook calculation mode Aspose.Cells | copy Excel table preserving references C# | Aspose.Cells range copy between worksheets

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, adds a new worksheet named "CopiedTable", defines a source range (A1:C10) on the original sheet, creates a matching destination range on the new sheet, copies the range while retaining all formulas, and saves the workbook, ensuring the original calculation mode remains unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Add a new worksheet to hold the copied table
            var newSheet = workbook.Worksheets.Add("CopiedTable");

            // Reference the source worksheet and the range containing the table with formulas
            var sourceSheet = workbook.Worksheets["Sheet1"]; // adjust name as needed
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Sheet1' not found.");
                return;
            }

            // Create source range using row/column indices (A1:C10 => rows 0-9, columns 0-2)
            var sourceRange = sourceSheet.Cells.CreateRange(0, 0, 10, 3);

            // Create a destination range on the new sheet with the same size as the source range
            var destRange = newSheet.Cells.CreateRange(0, 0, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy the source range to the destination range, preserving formulas
            sourceRange.Copy(destRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

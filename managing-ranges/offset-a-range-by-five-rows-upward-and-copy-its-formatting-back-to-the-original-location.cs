// Title: How to offset a cell range upward by five rows and copy its formatting back using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells offset a range five rows up and copy its style to the original range | Copy formatting from an offset range to the original cells with Aspose.Cells in .NET | Create range B10:D20 and apply style from range B5:D15 using Aspose.Cells | Aspose.Cells move a range upward and transfer cell styles in C# | Validate range boundaries when offsetting ranges in Aspose.Cells
// Tags: offset-range rows upward Aspose.Cells | copy-style between ranges Aspose.Cells | define-specific cell range Aspose.Cells | range-boundary validation Aspose.Cells | worksheet formatting transfer C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, defines the range B10:D20, calculates an offset range five rows above, copies the formatting from the offset range back to the original range, and saves the updated file.
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
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range (B10:D20)
            // B10 => row index 9, column index 1 (zero‑based)
            // 11 rows (10 through 20), 3 columns (B, C, D)
            AsposeRange originalRange = sheet.Cells.CreateRange(9, 1, 11, 3);

            // Calculate the offset range (5 rows upward)
            int offsetRow = originalRange.FirstRow - 5;
            if (offsetRow < 0)
            {
                Console.WriteLine("Offset range goes beyond the worksheet top boundary.");
                return;
            }

            AsposeRange offsetRange = sheet.Cells.CreateRange(offsetRow, originalRange.FirstColumn, originalRange.RowCount, originalRange.ColumnCount);

            // Copy the formatting from the offset range back to the original range
            originalRange.CopyStyle(offsetRange);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated formatting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

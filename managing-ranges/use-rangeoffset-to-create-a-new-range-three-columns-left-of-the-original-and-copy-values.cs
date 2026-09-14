// Title: Offset a range three columns to the left and copy its values using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to create a new range three columns left of an existing range and copy the original values into it. | Demonstrate how to check for a negative column index, offset a worksheet range by -3 columns, and duplicate the range data with the Aspose.Cells Range.Copy method.
// Common Searches: Aspose.Cells C# offset a range three columns left and copy data | How to duplicate an Excel range to another location with Aspose.Cells for .NET | Create a range with the same size at a different column index using Aspose.Cells | Prevent out‑of‑bounds column index when offsetting ranges in Aspose.Cells | Copy values from one worksheet range to another using Aspose.Cells Range.Copy
// Tags: range offset columns Aspose.Cells | copy range values Aspose.Cells C# | create range same dimensions Aspose.Cells | negative column index handling Aspose.Cells | excel workbook manipulation .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads 'input.xlsx', defines a source range (e.g., B2:D4), creates a new range three columns to the left (ensuring the column index stays non‑negative), copies the source values into the offset range, and saves the result as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range (example: B2:D4)
            AsposeRange originalRange = sheet.Cells.CreateRange("B2", "D4");

            // Calculate the offset position (three columns to the left)
            int offsetRow = originalRange.FirstRow;
            int offsetColumn = originalRange.FirstColumn - 3;
            if (offsetColumn < 0)
            {
                Console.WriteLine("Offset results in a negative column index. Operation aborted.");
                return;
            }

            // Create a new range at the offset position with the same size as the original range
            AsposeRange offsetRange = sheet.Cells.CreateRange(offsetRow, offsetColumn, originalRange.RowCount, originalRange.ColumnCount);

            // Copy the values from the original range to the offset range
            offsetRange.Copy(originalRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

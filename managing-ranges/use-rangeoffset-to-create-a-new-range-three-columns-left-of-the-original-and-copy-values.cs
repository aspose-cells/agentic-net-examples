// Title: Copy a cell range three columns to the left using Range.Offset in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy a range three columns left with Range.Offset | How to create an offset range in an Excel worksheet using Aspose.Cells | Shift cell values left by three columns programmatically with Aspose.Cells .NET | Duplicate range C3:E5 to a new location three columns earlier in C# | Using Aspose.Cells Range.Offset to copy values to a leftward offset range
// Tags: range offset Aspose.Cells C# | copy range values Excel .NET | create leftward offset range Aspose.Cells | shift Excel cells left programmatically | Aspose.Cells range manipulation example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, defines the source range C3:E5, calculates a new start column three positions to the left, creates an equally sized offset range, copies each cell value from the source to the offset range, and saves the result as output.xlsx.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range (example: C3:E5)
            AsposeRange originalRange = sheet.Cells.CreateRange("C3", "E5");

            // Calculate the start position for the offset range (three columns to the left)
            int newStartRow = originalRange.FirstRow;
            int newStartColumn = originalRange.FirstColumn - 3;

            // Ensure the new start column is not negative
            if (newStartColumn < 0)
            {
                Console.WriteLine("Offset results in an invalid column index.");
                return;
            }

            // Create the offset range with the same size as the original range
            AsposeRange offsetRange = sheet.Cells.CreateRange(
                newStartRow,
                newStartColumn,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Copy values from the original range to the offset range
            int rows = originalRange.RowCount;
            int cols = originalRange.ColumnCount;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    offsetRange[i, j].Value = originalRange[i, j].Value;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

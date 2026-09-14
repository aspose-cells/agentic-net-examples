// Title: Create an offset range from D4:F10 by moving it 3 rows down and 2 columns right using Aspose.Cells for .NET
// AI Prompts: Write C# code that takes an existing Aspose.Cells range and generates a new range shifted by a specific number of rows and columns while keeping the original size. | Show how to calculate the starting row and column for an offset range and instantiate it with Aspose.Cells API.
// Common Searches: how to offset a cell range in Aspose.Cells C# | Aspose.Cells move Excel range D4:F10 three rows down two columns right | create a new range with same dimensions at a different location using Aspose.Cells | C# example for shifting a range by rows and columns in Aspose.Cells | calculate start row and column for an offset range in Aspose.Cells
// Tags: offset range creation Aspose.Cells C# | shift range rows columns Aspose.Cells | preserve range dimensions Aspose.Cells | start row column calculation Aspose.Cells | move Excel range programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a workbook, defines the original range D4:F10, computes a new start position three rows lower and two columns to the right, builds an offset range with the same size, writes a value into its first cell, and saves the file as OffsetRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Original range D4:F10
            AsposeRange originalRange = worksheet.Cells.CreateRange("D4", "F10");

            // Offset the range 3 rows down and 2 columns right
            int rowOffset = 3;
            int columnOffset = 2;
            int startRow = originalRange.FirstRow + rowOffset;
            int startColumn = originalRange.FirstColumn + columnOffset;
            int rowCount = originalRange.RowCount;
            int columnCount = originalRange.ColumnCount;

            AsposeRange offsetRange = worksheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);

            // Example: put a value in the top‑left cell of the offset range
            offsetRange[0, 0].PutValue("Offset Start");

            // Define output file path
            string outputPath = "OffsetRange.xlsx";

            // Ensure the directory exists (handle null when outputPath has no directory part)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

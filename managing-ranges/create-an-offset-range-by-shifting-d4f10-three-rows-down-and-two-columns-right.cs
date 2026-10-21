// Title: Shift an existing range D4:F10 by three rows and two columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells how to move a range down three rows and right two columns in C# | C# create offset range from D4:F10 with Aspose.Cells | offsetting Excel cell range programmatically using Aspose.Cells .NET | shift range coordinates in an Aspose.Cells workbook example | Aspose.Cells create new range based on existing range start position
// Tags: Aspose.Cells offset range C# | create range from existing range Aspose.Cells | shift Excel cells programmatically .NET | range coordinate manipulation Aspose.Cells | Excel workbook range offset example

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample creates a new workbook, defines the original range D4:F10 on the first worksheet, then generates an offset range by adding three rows and two columns to the original start cell while preserving the original size, writes a value to the top‑left cell of the offset range, and saves the file as OffsetRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range D4:F10
            AsposeRange originalRange = sheet.Cells.CreateRange("D4", "F10");

            // Offset the range 3 rows down and 2 columns right
            int offsetRows = 3;
            int offsetColumns = 2;
            AsposeRange offsetRange = sheet.Cells.CreateRange(
                originalRange.FirstRow + offsetRows,
                originalRange.FirstColumn + offsetColumns,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Set a value in the top‑left cell of the offset range to verify the shift
            offsetRange[0, 0].PutValue("Offset Start");

            // Determine output file path
            string outputPath = "OffsetRangeDemo.xlsx";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

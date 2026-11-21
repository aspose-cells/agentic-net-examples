// Title: Use TableToRangeOptions.LastRow to retain formatting through row 15 when converting an Excel table to a range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert table to range keep formatting up to row 15 | TableToRangeOptions LastRow example preserving formatting in .NET | How to limit table-to-range conversion to specific rows using Aspose.Cells | C# Aspose.Cells keep cell styles when converting ListObject to range | Convert Excel ListObject to range while preserving first 15 rows formatting Aspose
// Tags: TableToRangeOptions LastRow | Aspose.Cells convert ListObject to range | preserve formatting when converting table | limit conversion rows with Aspose.Cells | C# Excel table to range conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads a workbook, checks for a table on the first worksheet, sets TableToRangeOptions.LastRow to 15 to keep formatting through that row, converts the table to a range with those options, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (sheet.ListObjects.Count == 0)
                throw new InvalidOperationException("No tables found in the first worksheet.");

            // Get the first table in the worksheet
            ListObject table = sheet.ListObjects[0];

            // Set options to keep formatting through row 15 (PreserveFormatting not available in this version)
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = 15 // 1‑based row number
            };

            // Convert the table to a range using the specified options
            table.ConvertToRange(options);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

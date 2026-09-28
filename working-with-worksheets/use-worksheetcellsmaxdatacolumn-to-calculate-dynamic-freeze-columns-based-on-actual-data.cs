// Title: Dynamically freeze all populated columns in an Excel worksheet using Aspose.Cells MaxDataColumn (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, determines the last column containing data via Worksheet.Cells.MaxDataColumn, and applies Worksheet.FreezePanes to lock those columns. | Show how to compute both MaxDataColumn and MaxDataRow and use them to set a dynamic freeze pane that freezes columns up to the rightmost data column while keeping rows unfrozen. | Create a reusable C# method named ApplyDynamicColumnFreeze that accepts a file path, calculates the maximum data column, and saves the workbook with the appropriate FreezePanes settings.
// Common Searches: Aspose.Cells C# freeze columns up to the last used column automatically | how to use MaxDataColumn to set dynamic freeze panes in Excel with .NET | C# code example for freezing all data columns in a worksheet using Aspose.Cells | determine rightmost data column and apply FreezePanes in Aspose.Cells for .NET | dynamic column freeze based on worksheet content Aspose.Cells tutorial
// Tags: Worksheet.FreezePanes column locking Aspose.Cells | max data column retrieval C# Aspose.Cells | auto column freeze based on data .NET | last used column calculation Aspose.Cells | dynamic column freeze implementation C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, uses Worksheet.Cells.MaxDataColumn and MaxDataRow to find the furthest populated column and row, freezes all columns up to that column with Worksheet.FreezePanes, and saves the modified file.
class DynamicFreezeColumns
{
    static void Main()
    {
        string inputPath = @"C:\Input\Sample.xlsx";
        string outputPath = @"C:\Output\Sample_Frozen.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the rightmost column that contains data (zero‑based index)
            int maxDataColumn = sheet.Cells.MaxDataColumn;

            // Determine the bottommost row that contains data (zero‑based index)
            int maxDataRow = sheet.Cells.MaxDataRow;

            // Freeze all columns that contain data.
            // FreezePanes(row, column, totalRows, totalColumns) freezes rows above 'row' and columns left of 'column'.
            // No rows are frozen (row = 0), columns left of (maxDataColumn + 1) are frozen.
            sheet.FreezePanes(0, maxDataColumn + 1, maxDataRow + 1, maxDataColumn + 1);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

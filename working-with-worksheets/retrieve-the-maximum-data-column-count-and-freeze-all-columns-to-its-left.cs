// Title: Freeze columns to the left of the last populated column in an Excel sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, finds the worksheet's MaxDataColumn, and calls FreezePanes to lock all columns before that column. | Show how to combine Cells.MaxDataColumn and FreezePanes to dynamically freeze columns up to the used range in a .NET Excel workbook. | Create a .NET example that determines the last data column, freezes the pane at column (maxDataColumn + 1), and saves the modified workbook.
// Common Searches: aspocells c# freeze panes up to the last used column | how to lock columns before the max data column in an Excel workbook using Aspose.Cells | retrieve maximum data column and apply FreezePanes in .NET | dynamic freeze panes based on used range with Aspose.Cells for .NET | C# example for freezing columns left of the populated area in Excel
// Tags: Aspose.Cells FreezePanes using MaxDataColumn | C# get last populated column Excel | dynamic column freeze Aspose.Cells | Excel used range column count .NET | lock columns left of used range Aspose.Cells | determine MaxDataColumn property C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'input.xlsx', accesses the first worksheet, uses Cells.MaxDataColumn (and MaxDataRow) to locate the furthest populated column, then freezes all columns to its left with FreezePanes, and saves the result as 'output.xlsx' while handling errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Work with the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Determine the maximum column and row that contain data (0‑based)
            int maxDataColumn = worksheet.Cells.MaxDataColumn;
            int maxDataRow = worksheet.Cells.MaxDataRow;

            // Freeze all columns to the left of the max data column.
            // row = 0 (no rows frozen), column = maxDataColumn + 1.
            // totalRows and totalColumns define the scrollable area size.
            worksheet.FreezePanes(0, maxDataColumn + 1, maxDataRow + 1, maxDataColumn + 1);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

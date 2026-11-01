// Title: Freeze the first ten columns in worksheets that have more than ten columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze first ten columns only when worksheet exceeds ten columns | conditionally apply FreezePanes based on column count in Aspose.Cells | detect populated column count and set freeze pane in Excel workbook with Aspose.Cells | C# code to freeze panes for sheets with more than 10 columns using Aspose.Cells | dynamic FreezePanes implementation for Excel worksheets in .NET
// Tags: conditional FreezePanes Aspose.Cells .NET | freeze first ten columns Excel workbook C# | max data column detection Aspose.Cells | worksheet iteration for freeze pane Aspose.Cells | Excel column count based freeze pane C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, iterates through each worksheet, determines the number of populated columns, and if a sheet contains more than ten columns, it freezes the first ten columns (without freezing rows) using the FreezePanes method, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the number of columns that contain data (zero‑based index + 1)
                int columnCount = sheet.Cells.MaxDataColumn + 1;

                // Apply freeze pane only if the worksheet has more than ten columns
                if (columnCount > 10)
                {
                    // Determine total rows and columns for the FreezePanes method
                    int totalRows = sheet.Cells.MaxDataRow + 1;
                    int totalColumns = columnCount; // same as sheet.Cells.MaxDataColumn + 1

                    // Freeze the first ten columns (no rows are frozen)
                    // Parameters: totalRows, totalColumns, frozenRows, frozenColumns
                    sheet.FreezePanes(totalRows, totalColumns, 0, 10);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

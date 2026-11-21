// Title: Freeze columns up to the last populated column in an Excel worksheet using Worksheet.Cells.MaxDataColumn with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells freeze panes up to the last data column | How to use Worksheet.Cells.MaxDataColumn to set FreezePanes in .NET | Dynamic column freeze based on worksheet content with Aspose.Cells | Determine the last populated column and apply FreezePanes in C# | Automatically freeze rows and columns according to data range using Aspose.Cells
// Tags: freeze panes using Worksheet.Cells.MaxDataColumn | determine last populated column Aspose.Cells | dynamic column freeze Aspose.Cells .NET | apply FreezePanes based on data range C# | max data column calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads an existing workbook, uses Worksheet.Cells.MaxDataColumn to find the last column containing data, optionally gets the total rows, then calls FreezePanes to lock all columns up to that column and saves the updated file.
    class DynamicFreezeColumns
    {
        static void Main()
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Find the last column that contains data (0‑based index)
                int maxDataColumn = sheet.Cells.MaxDataColumn;

                // Determine total rows to include in the freeze pane (optional)
                int totalRows = sheet.Cells.MaxDataRow + 1;

                // Freeze all columns up to the last data column.
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(0, maxDataColumn + 1, totalRows, maxDataColumn + 1);

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
}

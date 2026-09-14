// Title: Apply a custom number format to column C while skipping the first two rows with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, creates a range for column C starting at row 3, and applies the custom number format '#,##0.00;[Red]-#,##0.00' using Aspose.Cells. | Provide a reusable method that takes a worksheet, column index, start‑row offset, and format string, then creates the appropriate range and applies the style with Aspose.Cells. | Show how to modify the example to format column D instead of C and use a different custom format while keeping the two‑row offset.
// Common Searches: Aspose.Cells how to format a whole column with a custom number format after skipping header rows | C# create a range with offset rows and apply StyleFlag.All in Aspose.Cells | Apply red negative number formatting to column C using Aspose.Cells .NET | Set custom numeric format for dynamic range based on last data row Aspose.Cells | Skip first two rows when applying number format to Excel column with Aspose.Cells
// Tags: apply custom number format Aspose.Cells | create range with start row offset Aspose.Cells | format column C Aspose.Cells .NET | use StyleFlag.All Aspose.Cells | skip header rows number formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing workbook, creates a range covering column C from row 3 to the last data row, applies a custom number format '#,##0.00;[Red]-#,##0.00' using a style with StyleFlag.All, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Determine the last row that contains data
                int lastDataRow = worksheet.Cells.MaxDataRow;

                // Define the offset (skip the first two rows)
                int startRow = 2;               // Zero‑based index: row 3 in Excel
                int startColumn = 2;            // Column C (zero‑based)

                // Calculate how many rows to include in the range
                int totalRows = (lastDataRow >= startRow) ? (lastDataRow - startRow + 1) : 0;

                // Apply the custom number format only if there are rows to format
                if (totalRows > 0)
                {
                    // Create a range that covers column C from the offset row to the last data row
                    var range = worksheet.Cells.CreateRange(startRow, startColumn, totalRows, 1);

                    // Create a style with the desired custom number format
                    var style = workbook.CreateStyle();
                    style.Custom = "#,##0.00;[Red]-#,##0.00";   // Example custom format

                    // Apply the style to the range (apply all style attributes)
                    var flag = new StyleFlag { All = true };
                    range.ApplyStyle(style, flag);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

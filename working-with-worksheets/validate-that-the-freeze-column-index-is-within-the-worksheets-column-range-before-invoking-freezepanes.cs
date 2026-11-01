// Title: How to validate a freeze column index against a worksheet’s column range before using FreezePanes in Aspose.Cells for .NET
// AI Prompts: Write C# code that checks if a zero‑based column index is within Worksheet.Cells.MaxColumn and throws an ArgumentOutOfRangeException before calling Worksheet.FreezePanes. | Generate a method that accepts a workbook path and a freeze column index, validates the index against the worksheet’s column count, applies FreezePanes to freeze the specified number of columns, and saves the workbook. | Create error‑handling logic that logs a clear message when the freeze column index exceeds the worksheet’s column range in an Aspose.Cells application.
// Common Searches: Aspose.Cells C# verify column index before FreezePanes to avoid ArgumentOutOfRangeException | how to check max column in worksheet when freezing panes with Aspose.Cells | prevent out‑of‑range errors when freezing first N columns in an Excel file using Aspose.Cells .NET
// Tags: validate freeze column index Aspose.Cells | Worksheet.Cells.MaxColumn range check | FreezePanes column bounds validation | C# Aspose.Cells freeze panes safety | Excel worksheet column range verification

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, uses Worksheet.Cells.MaxColumn to ensure the requested freeze column index is within the worksheet’s column range, throws an ArgumentOutOfRangeException if it is not, applies FreezePanes to freeze the desired number of columns, and saves the updated file.
class FreezePaneExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze column index (zero‑based). Example: freeze first 3 columns.
            int freezeColumnIndex = 3;

            // Validate that the freeze column index is within the worksheet's column range.
            // Cells.MaxColumn returns the zero‑based index of the last column that contains data.
            int maxColumnIndex = worksheet.Cells.MaxColumn; // zero‑based
            if (freezeColumnIndex < 0 || freezeColumnIndex > maxColumnIndex)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(freezeColumnIndex),
                    $"Freeze column index ({freezeColumnIndex}) is outside the worksheet's column range (0‑{maxColumnIndex}).");
            }

            // Determine the total rows and columns for the pane (required by the 4‑parameter overload)
            int totalRows = worksheet.Cells.MaxRow + 1;      // MaxRow is zero‑based
            int totalColumns = worksheet.Cells.MaxColumn + 1;

            // Apply freeze panes: freeze columns to the left of freezeColumnIndex.
            // The overload FreezePanes(row, column, totalRows, totalColumns) freezes rows above 'row' and columns left of 'column'.
            worksheet.FreezePanes(0, freezeColumnIndex, totalRows, totalColumns);

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

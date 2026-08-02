// Title: Validate Freeze Column Index Before Using Worksheet.FreezePanes in Aspose.Cells for .NET
// Description: Demonstrates how to retrieve the workbook's maximum column index with workbook.Settings.MaxColumn, verify that a zero‑based freezeColumnIndex is within the valid range, and call Worksheet.FreezePanes only when the index is safe, then save the workbook.
// Keywords: Aspose.Cells | C# FreezePanes validation | max column index | worksheet freeze column range | prevent FreezePanes error | Excel column limit Aspose | .NET spreadsheet API
// Common Searches: Aspose.Cells check column index before FreezePanes | maximum column count in Aspose.Cells workbook | how to avoid out‑of‑range error when freezing panes | C# example for validating FreezePanes parameters | worksheet.FreezePanes column range validation
// Developer Intent: Confirm that the column index supplied to Worksheet.FreezePanes does not exceed the workbook's allowed column range.
// Use Cases: Guard against runtime exceptions when user‑provided column numbers are out of bounds. | Dynamically freeze the first N columns only if N is less than or equal to the format's column limit. | Create a reusable helper that logs an error and skips FreezePanes when the index is invalid.
// AI Prompts: Generate a C# method that takes a column index, checks it against workbook.Settings.MaxColumn, applies Worksheet.FreezePanes if valid, and returns a success flag. | Write code that logs a detailed warning and continues execution without calling FreezePanes when the requested column exceeds the maximum for the current Excel version using Aspose.Cells. | Provide a user‑input driven example that validates the freeze column range, freezes panes on success, and handles invalid input gracefully.

using System;
using Aspose.Cells;

namespace FreezePanesValidationDemo
{
    // Demonstrates how to retrieve the workbook's maximum column index with workbook.Settings.MaxColumn, verify that a zero‑based freezeColumnIndex is within the valid range, and call Worksheet.FreezePanes only when the index is safe, then save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze column index (zero‑based). Change as needed.
            int freezeColumnIndex = 300;

            // Retrieve the maximum allowed column index for the workbook format
            int maxColumnIndex = workbook.Settings.MaxColumn; // zero‑based

            // Validate that the freeze column is within the worksheet's column range
            if (freezeColumnIndex >= 0 && freezeColumnIndex <= maxColumnIndex)
            {
                // Freeze panes at row 0, column freezeColumnIndex
                // The last two parameters specify the number of frozen rows and columns
                worksheet.FreezePanes(0, freezeColumnIndex, 0, freezeColumnIndex);
                Console.WriteLine($"FreezePanes applied at column index {freezeColumnIndex}.");
            }
            else
            {
                Console.WriteLine($"Error: Column index {freezeColumnIndex} is outside the valid range (0‑{maxColumnIndex}).");
            }

            // Save the workbook (uses the provided save rule)
            workbook.Save("ValidatedFreezePanes.xlsx");
        }
    }
}

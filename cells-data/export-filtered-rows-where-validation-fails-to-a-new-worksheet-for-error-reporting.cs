// Title: Export rows that fail custom validation to a separate worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy rows that do not meet validation rules to another sheet in C# | C# generate error report worksheet for invalid Excel rows using Aspose.Cells | How to filter Excel rows by custom validation and export them to a new tab with Aspose.Cells | Save failed validation rows to a separate worksheet in a .NET Excel processing app | Aspose.Cells example for exporting invalid data rows to an error sheet
// Tags: Aspose.Cells CopyRows method | export invalid rows to separate sheet | custom row validation with Aspose.Cells | error report worksheet generation | C# Excel validation and export using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input workbook, adds an "ErrorReport" worksheet, copies the header row, validates each data row (e.g., non‑empty column A and positive numeric column B), copies rows that fail the validation to the error sheet using Cells.CopyRows, and saves the workbook with the error report.
class ExportValidationErrors
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet for error reporting
            Worksheet errorSheet = workbook.Worksheets.Add("ErrorReport");

            // Copy the header row (assumed to be the first row) to the error sheet
            // Use the overload that accepts source Cells as the first argument
            errorSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 1);

            // Index in the error sheet where the next failed row will be written
            int errorRowIndex = 1; // start after header

            // Determine the range of used rows in the source sheet
            int maxRow = sourceSheet.Cells.MaxDataRow;

            // Iterate through each data row (skip header row at index 0)
            for (int row = 1; row <= maxRow; row++)
            {
                // Perform validation on the current row
                if (!ValidateRow(sourceSheet, row))
                {
                    // Copy the entire row that failed validation to the error sheet
                    errorSheet.Cells.CopyRows(sourceSheet.Cells, row, errorRowIndex, 1);
                    errorRowIndex++;
                }
            }

            // Save the workbook with the error report sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="sheet">Worksheet containing the data.</param>
    /// <param name="rowIndex">Zero‑based index of the row to validate.</param>
    /// <returns>True if the row passes validation; otherwise false.</returns>
    private static bool ValidateRow(Worksheet sheet, int rowIndex)
    {
        // Example validation:
        // - Column A must be a non‑empty string.
        // - Column B must be a numeric value greater than 0.
        // Adjust column indexes (0‑based) as needed for your data.

        // Validate Column A (string not empty)
        string colA = sheet.Cells[rowIndex, 0].StringValue?.Trim();
        if (string.IsNullOrEmpty(colA))
            return false;

        // Validate Column B (numeric > 0)
        if (!double.TryParse(sheet.Cells[rowIndex, 1].StringValue, out double colB) || colB <= 0)
            return false;

        // Add additional validation rules here...

        // If all checks pass, the row is valid
        return true;
    }
}

// Title: Set a custom calculation error string for an Excel PivotTable using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set custom error string for PivotTable calculation errors | display custom error message in Excel PivotTable with Aspose.Cells | enable DisplayErrorString and assign ErrorString in a PivotTable using .NET | handle PivotTable calculation errors by showing a custom message in C# | Aspose.Cells example to show custom error text for a PivotTable
// Tags: Aspose.Cells PivotTable DisplayErrorString | set PivotTable ErrorString C# | custom calculation error message Excel PivotTable | C# Aspose.Cells modify PivotTable error handling | Excel workbook save with custom PivotTable error text

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, checks for a PivotTable on the first worksheet, enables the DisplayErrorString flag, assigns a custom message to the ErrorString property, and saves the workbook. It also includes handling for missing files, password‑protected workbooks, and other Aspose.Cells exceptions.
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
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (no password supplied)
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // If a pivot table exists, set custom error handling
            if (worksheet.PivotTables.Count > 0)
            {
                var pivotTable = worksheet.PivotTables[0];
                pivotTable.DisplayErrorString = true;
                pivotTable.ErrorString = "Calculation error occurred. Please verify the source data.";
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files or other Cells-specific errors
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Provide a password to load it.");
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

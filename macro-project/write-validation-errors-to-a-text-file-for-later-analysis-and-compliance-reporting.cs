// Title: Create a C# console program that validates specific cells in an Excel workbook with Aspose.Cells and writes validation errors to a text file
// AI Prompts: Load an .xlsx workbook with Aspose.Cells in C#, verify a predefined list of required cells for empty values, perform numeric range checks on designated cells, collect any failures, and save the error messages to a .txt report. | Refactor the validation utility to read required‑cell and range‑check definitions from a JSON configuration file, then generate the validation results as a CSV log. | Update the code so that each run appends new validation messages to an existing log file instead of overwriting the previous contents.
// Common Searches: c# Aspose.Cells how to check specific cells for emptiness and log errors to a text file | validate numeric range of Excel cells with Aspose.Cells and output a validation report | write Excel validation results to .txt using Aspose.Cells in a .NET console application | Aspose.Cells required cell validation example C# | generate compliance error log from workbook with Aspose.Cells
// Tags: Aspose.Cells required‑cell validation C# | Aspose.Cells numeric range check .NET | write validation errors to text file Aspose.Cells | Excel workbook cell validation using Aspose.Cells | C# console error reporting for Excel with Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook via Aspose.Cells, checks designated required cells for missing or empty values, validates numeric ranges for specific cells, aggregates any error messages, and writes the list to a text file (or creates an empty file when no errors are found).
class ValidationErrorReporter
{
    static void Main(string[] args)
    {
        // Path to the Excel file to validate
        string excelPath = "input.xlsx";
        // Path to the text file where validation errors will be written
        string errorReportPath = "validation_errors.txt";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Collect validation errors
        List<string> errors = new List<string>();

        // Example validation: check that required cells are not empty
        // Define required cells per worksheet (sheet name -> list of cell names)
        var requiredCells = new Dictionary<string, List<string>>
        {
            { "Sheet1", new List<string> { "A1", "B2", "C3" } },
            { "Sheet2", new List<string> { "D4", "E5" } }
        };

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (!requiredCells.ContainsKey(sheet.Name))
                continue;

            foreach (string cellName in requiredCells[sheet.Name])
            {
                Cell cell = sheet.Cells[cellName];
                if (cell == null || cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    errors.Add($"Worksheet '{sheet.Name}': Cell '{cellName}' is empty or missing.");
                }
            }
        }

        // Additional example validation: numeric range check for specific cells
        // Define range checks (sheet name -> cell name -> (min, max))
        var rangeChecks = new Dictionary<string, Dictionary<string, (double min, double max)>>()
        {
            {
                "Sheet1", new Dictionary<string, (double, double)>
                {
                    { "D5", (0, 100) },
                    { "E6", (10, 50) }
                }
            }
        };

        foreach (var sheetEntry in rangeChecks)
        {
            Worksheet sheet = workbook.Worksheets[sheetEntry.Key];
            foreach (var cellEntry in sheetEntry.Value)
            {
                Cell cell = sheet.Cells[cellEntry.Key];
                if (cell == null || !double.TryParse(cell.StringValue, out double value))
                {
                    errors.Add($"Worksheet '{sheet.Name}': Cell '{cellEntry.Key}' does not contain a valid number.");
                }
                else
                {
                    var (min, max) = cellEntry.Value;
                    if (value < min || value > max)
                    {
                        errors.Add($"Worksheet '{sheet.Name}': Cell '{cellEntry.Key}' value {value} is outside the allowed range [{min}, {max}].");
                    }
                }
            }
        }

        // Write errors to the text file
        if (errors.Count > 0)
        {
            File.WriteAllLines(errorReportPath, errors);
            Console.WriteLine($"Validation completed. {errors.Count} error(s) written to '{errorReportPath}'.");
        }
        else
        {
            // Ensure the file exists but is empty if no errors
            File.WriteAllText(errorReportPath, string.Empty);
            Console.WriteLine("Validation completed. No errors found.");
        }
    }
}

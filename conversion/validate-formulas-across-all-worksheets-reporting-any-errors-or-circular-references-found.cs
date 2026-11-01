// Title: Validate all formulas in an Excel workbook and detect circular references using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to find formula errors and circular references in an Excel file | validate formulas across all worksheets with Aspose.Cells .NET API | detect and report Excel formula calculation exceptions using Aspose.Cells in C# | iterate through cells that contain formulas and capture errors with Aspose.Cells for .NET | Aspose.Cells calculate each formula and log issues in a workbook
// Tags: Aspose.Cells formula validation | detect circular references Aspose.Cells | calculate cell formulas .NET | Excel workbook formula error reporting C# | iterate worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, walks through every worksheet and each cell that contains a formula, attempts to calculate the formula with Aspose.Cells' CalculationOptions, captures any exceptions (including circular‑reference errors), records the sheet name, cell address and error message, and finally prints a concise report of all formula issues found.
class FormulaValidator
{
    static void Main(string[] args)
    {
        // Ensure an input file path is provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: FormulaValidator <input-file>");
            return;
        }

        string inputPath = args[0];

        // Verify that the file exists before attempting to load
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Collect validation issues
            var issues = new List<(string SheetName, string CellName, string ErrorMessage)>();

            // Iterate through all worksheets and cells that contain formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        try
                        {
                            // Attempt to calculate the formula using default options
                            cell.Calculate(calcOptions);
                        }
                        catch (Exception ex)
                        {
                            // Record the error details
                            issues.Add((sheet.Name, cell.Name, ex.Message));
                        }
                    }
                }
            }

            // Report validation results
            if (issues.Count == 0)
            {
                Console.WriteLine("No formula errors or circular references were found.");
            }
            else
            {
                Console.WriteLine("Formula validation issues detected:");
                foreach (var issue in issues)
                {
                    Console.WriteLine($"Worksheet: {issue.SheetName}, Cell: {issue.CellName}, Error: {issue.ErrorMessage}");
                }
            }

            // Optional: save the workbook if any modifications are needed
            // workbook.Save("validated_output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

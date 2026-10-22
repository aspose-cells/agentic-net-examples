// Title: Create a .NET report of Excel formulas that use unsupported functions with Aspose.Cells
// Common Searches: aspnet cells generate report of formulas containing unsupported functions | c# find Excel formulas that use functions not in supported list using Aspose.Cells | how to extract and list unknown functions from Excel cell formulas in .NET | regex to identify unsupported Excel functions in workbook with Aspose.Cells | create audit workbook of invalid formulas in Excel using Aspose.Cells C#
// Tags: unsupported Excel function detection Aspose.Cells | formula audit report generation .NET | extract function names from Excel formulas C# | cross‑reference supported function list Aspose.Cells | programmatic Excel workbook report creation | regex based formula parsing Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an input .xlsx file with Aspose.Cells, defines a HashSet of supported Excel functions, scans every worksheet and cell, uses a regular expression to pull function names from each formula, records any functions not in the supported set, and writes an output workbook that lists the sheet, cell address, original formula, and the unsupported functions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string reportPath = "UnsupportedFormulasReport.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook sourceWorkbook;
        try
        {
            // Load the source workbook (no password supplied)
            sourceWorkbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Define a set of supported Excel functions (add more as needed)
        var supportedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "IF", "AND", "OR", "NOT",
            "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "LEFT", "RIGHT", "MID",
            "LEN", "ROUND", "ROUNDUP", "ROUNDDOWN", "CONCATENATE", "TEXT"
        };

        // Create a new workbook for the report
        var reportWorkbook = new Workbook();
        var reportSheet = reportWorkbook.Worksheets[0];
        reportSheet.Name = "Unsupported Formulas";

        // Write header row
        reportSheet.Cells[0, 0].PutValue("Sheet");
        reportSheet.Cells[0, 1].PutValue("Cell");
        reportSheet.Cells[0, 2].PutValue("Formula");
        reportSheet.Cells[0, 3].PutValue("Unsupported Functions");

        int reportRow = 1;

        // Regex to extract function names from a formula (e.g., SUM( )
        var functionRegex = new Regex(@"([A-Za-z_][A-Za-z0-9_]*)\s*\(", RegexOptions.Compiled);

        // Iterate through all worksheets and cells in the source workbook
        foreach (Worksheet ws in sourceWorkbook.Worksheets)
        {
            Cells cells = ws.Cells;
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Find all function names used in the formula
                    MatchCollection matches = functionRegex.Matches(cell.Formula);
                    var unsupported = new List<string>();

                    foreach (Match match in matches)
                    {
                        string funcName = match.Groups[1].Value;
                        if (!supportedFunctions.Contains(funcName) && !unsupported.Contains(funcName))
                        {
                            unsupported.Add(funcName);
                        }
                    }

                    // If any unsupported functions are found, add an entry to the report
                    if (unsupported.Count > 0)
                    {
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(cell.Name);
                        reportSheet.Cells[reportRow, 2].PutValue(cell.Formula);
                        reportSheet.Cells[reportRow, 3].PutValue(string.Join(", ", unsupported));
                        reportRow++;
                    }
                }
            }
        }

        // Adjust column widths for readability
        reportSheet.AutoFitColumns();

        // Save the report workbook with error handling
        try
        {
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved successfully to \"{reportPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving report: {ex.Message}");
        }
    }
}

// Title: Create a C# Aspose.Cells utility to report Excel formulas that use unsupported functions
// Common Searches: C# Aspose.Cells find formulas with functions not supported by the library | how to generate a report of unsupported Excel functions using Aspose.Cells .NET | scan workbook for unsupported formula functions with Aspose.Cells C# example | extract unsupported function names from Excel formulas in a .NET application | list cells containing unsupported Excel functions using Aspose.Cells API
// Tags: Aspose.Cells unsupported function detection | C# generate Excel formula report | regex function extraction Aspose.Cells | Excel workbook formula analysis .NET | unsupported Excel functions report generation

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The sample loads an input workbook, defines a list of functions that Aspose.Cells supports, iterates through every worksheet and cell, extracts function names from each formula using a regular expression, records any formula that contains a function not in the supported list, and writes a new Excel file that lists the sheet name, cell address, full formula, and the unsupported function.
class UnsupportedFunctionsReport
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string reportPath = "UnsupportedFunctionsReport.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook srcWorkbook;
        try
        {
            srcWorkbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // List to store information about unsupported formulas
        var unsupportedFormulas = new List<(string SheetName, string CellName, string Formula, string Function)>();

        // Simple set of functions known to be supported by Aspose.Cells (extend as needed)
        var supportedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SUM", "AVERAGE", "MAX", "MIN", "COUNT", "IF", "AND", "OR", "NOT",
            "LEN", "LEFT", "RIGHT", "MID", "VLOOKUP", "HLOOKUP", "INDEX",
            "MATCH", "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "POWER"
        };

        // Regex to extract function names from a formula (e.g., SUM(A1:B2) -> SUM)
        Regex funcRegex = new Regex(@"([A-Z][A-Z0-9_]*)\s*\(", RegexOptions.IgnoreCase);

        // Iterate through each worksheet
        foreach (Worksheet sheet in srcWorkbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through each cell in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Find all function names used in the formula
                    foreach (Match match in funcRegex.Matches(formula))
                    {
                        string funcName = match.Groups[1].Value;

                        // Determine if the function is unsupported
                        if (!supportedFunctions.Contains(funcName))
                        {
                            unsupportedFormulas.Add((sheet.Name, cell.Name, formula, funcName));
                            break; // stop after first unsupported function per cell
                        }
                    }
                }
            }
        }

        // Create a new workbook for the report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];
        Cells reportCells = reportSheet.Cells;

        // Write header row
        reportCells["A1"].PutValue("Sheet");
        reportCells["B1"].PutValue("Cell");
        reportCells["C1"].PutValue("Formula");
        reportCells["D1"].PutValue("Unsupported Function");

        // Populate report rows with collected data
        int rowIndex = 1; // zero‑based index; row 1 is the second row in the sheet
        foreach (var entry in unsupportedFormulas)
        {
            reportCells[rowIndex, 0].PutValue(entry.SheetName);
            reportCells[rowIndex, 1].PutValue(entry.CellName);
            reportCells[rowIndex, 2].PutValue(entry.Formula);
            reportCells[rowIndex, 3].PutValue(entry.Function);
            rowIndex++;
        }

        // Adjust column widths for readability
        reportSheet.AutoFitColumns();

        // Save the report workbook
        try
        {
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to \"{reportPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save report: {ex.Message}");
        }
    }
}

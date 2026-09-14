// Title: C# Aspose.Cells program to generate a report of workbook formulas that use unsupported Excel functions
// AI Prompts: Write C# code with Aspose.Cells that iterates every worksheet, extracts each formula, compares the called functions against a predefined whitelist, and creates a new Excel file listing the sheet name, cell address, full formula, and the unsupported functions. | Adapt the script to load the whitelist of supported functions from an external text or JSON file and include the complete argument list of each unsupported function in the generated report. | Enhance the output workbook by applying conditional formatting that highlights rows containing more than one unsupported function.
// Common Searches: how to find Excel functions not supported by Aspose.Cells in a .NET workbook | C# Aspose.Cells generate list of cells with unknown functions | scan an Excel file for non‑standard formulas using Aspose.Cells | export unsupported formula functions to a separate workbook with Aspose.Cells | create compliance report for Excel functions in a .NET application
// Tags: unsupported function detection Aspose.Cells | formula audit workbook .NET | extract Excel function names C# | generate unsupported functions report Excel | conditional formatting rows multiple unsupported functions

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

// The example loads an input workbook, defines a HashSet of supported Excel function names, walks through every cell that contains a formula, uses a regular expression to capture each function call, records any functions not present in the supported set, and writes the sheet name, cell address, original formula, and the list of unsupported functions to a new workbook saved as UnsupportedFunctionsReport.xlsx.
class UnsupportedFunctionsReport
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook srcWorkbook = new Workbook("input.xlsx");

        // Define a set of supported Excel functions (add more as needed)
        HashSet<string> supportedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SUM","AVERAGE","MIN","MAX","COUNT","IF","AND","OR","NOT",
            "VLOOKUP","HLOOKUP","INDEX","MATCH","LEFT","RIGHT","MID",
            "LEN","ROUND","ROUNDUP","ROUNDDOWN","CONCATENATE","TEXT"
        };

        // List to store information about formulas that use unsupported functions
        List<(string SheetName, string CellName, string Formula, List<string> UnsupportedFuncs)> unsupportedFormulas
            = new List<(string, string, string, List<string>)>();

        // Regex to extract function names from a formula (e.g., SUM(A1:B2) -> SUM)
        Regex funcRegex = new Regex(@"([A-Z][A-Z0-9\.]*)\s*\(", RegexOptions.Compiled);

        // Iterate through each worksheet and each cell that contains a formula
        foreach (Worksheet sheet in srcWorkbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string formula = cell.Formula;
                    MatchCollection matches = funcRegex.Matches(formula);
                    List<string> unsupported = new List<string>();

                    foreach (Match match in matches)
                    {
                        string funcName = match.Groups[1].Value;
                        // If the function is not in the supported list, record it
                        if (!supportedFunctions.Contains(funcName))
                        {
                            unsupported.Add(funcName);
                        }
                    }

                    if (unsupported.Count > 0)
                    {
                        unsupportedFormulas.Add((sheet.Name, cell.Name, formula, unsupported));
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
        reportCells["D1"].PutValue("Unsupported Functions");

        // Populate the report with collected data
        int row = 1; // zero‑based index; row 1 is the second row in the sheet
        foreach (var entry in unsupportedFormulas)
        {
            reportCells[row, 0].PutValue(entry.SheetName);
            reportCells[row, 1].PutValue(entry.CellName);
            reportCells[row, 2].PutValue(entry.Formula);
            reportCells[row, 3].PutValue(string.Join(", ", entry.UnsupportedFuncs));
            row++;
        }

        // Adjust column widths for readability
        reportSheet.AutoFitColumns();

        // Save the report workbook
        reportWorkbook.Save("UnsupportedFunctionsReport.xlsx");
    }
}

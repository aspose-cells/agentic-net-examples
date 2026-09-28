// Title: Export a summary worksheet that aggregates the count of each formula function per worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet, extracts the function name from each formula via a regular expression, tallies the occurrences per function per sheet, and creates a new 'Summary' sheet listing Worksheet, Function, and Count. | Extend the solution to ignore hidden worksheets and handle array formulas when counting functions, then save the updated workbook.
// Common Searches: aspnet count Excel formula functions per sheet using Aspose.Cells | C# generate formula usage summary worksheet with Aspose.Cells | extract Excel function name from formula regex Aspose.Cells C# | how to create a summary sheet that lists function counts across worksheets in .NET | skip hidden worksheets when summarizing formulas Aspose.Cells
// Tags: Aspose.Cells count formula functions per worksheet | C# regex extract Excel function name Aspose.Cells | add summary sheet with function usage statistics | ignore hidden worksheets Aspose.Cells | array formulas handling in Aspose.Cells function tally

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, iterates each visible worksheet, uses a regular expression to capture the function name of every formula cell, aggregates the count of each function per sheet, creates a new 'Summary' worksheet with columns Worksheet, Function, and Count, auto‑fits the columns, and saves the workbook.
class FormulaSummaryExporter
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to hold counts: Worksheet -> (Function -> Count)
        var summary = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        // Regular expression to extract the function name from a formula (e.g., =SUM(A1:B2) -> SUM)
        Regex funcRegex = new Regex(@"^=([A-Za-z_][A-Za-z0-9_]*)\s*\(", RegexOptions.Compiled);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Skip a previously existing summary sheet (if any)
            if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                continue;

            var funcCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Cells cells = sheet.Cells;

            // Iterate over all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!cell.IsFormula)
                    continue;

                string formula = cell.Formula;
                if (string.IsNullOrEmpty(formula))
                    continue;

                Match match = funcRegex.Match(formula);
                if (match.Success)
                {
                    string funcName = match.Groups[1].Value.ToUpperInvariant();

                    if (funcCounts.ContainsKey(funcName))
                        funcCounts[funcName]++;
                    else
                        funcCounts[funcName] = 1;
                }
            }

            // Store counts for the current worksheet
            summary[sheet.Name] = funcCounts;
        }

        // Add a new worksheet for the summary
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[summaryIndex];
        summarySheet.Name = "Summary";

        // Write headers
        summarySheet.Cells["A1"].PutValue("Worksheet");
        summarySheet.Cells["B1"].PutValue("Function");
        summarySheet.Cells["C1"].PutValue("Count");

        int row = 2; // Start writing data from the second row

        // Populate the summary data
        foreach (var sheetEntry in summary)
        {
            string sheetName = sheetEntry.Key;
            var funcDict = sheetEntry.Value;

            foreach (var funcEntry in funcDict)
            {
                summarySheet.Cells[row, 0].PutValue(sheetName);          // Column A
                summarySheet.Cells[row, 1].PutValue(funcEntry.Key);    // Column B
                summarySheet.Cells[row, 2].PutValue(funcEntry.Value); // Column C
                row++;
            }
        }

        // Auto-fit columns for better readability
        summarySheet.AutoFitColumns();

        // Save the workbook with the new summary sheet
        workbook.Save("output.xlsx");
    }
}

// Title: Create a Summary Sheet with Formula Counts by Function Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, skips any existing "Summary" tab, scans each worksheet for formula cells, extracts the function name, tallies occurrences per sheet, adds a new "Summary" worksheet with columns for worksheet name, function, and count, auto‑fits columns, and saves the updated file.
// Keywords: Aspose.Cells formula count | C# Excel function summary | aggregate formula usage | add summary worksheet .NET | Excel function statistics Aspose | count formulas per sheet | extract function name from formula
// Common Searches: how to count Excel functions per worksheet using Aspose.Cells | create a summary tab that lists formula usage in C# | Aspose.Cells enumerate formula cells and group by function | generate formula statistics workbook Aspose .NET | add summary sheet with function counts in Excel
// Developer Intent: Generate a new worksheet that lists, for every existing sheet, how many times each formula function appears.
// Use Cases: Audit workbook to identify the most used functions on each sheet. | Produce documentation showing formula distribution across worksheets. | Spot sheets that heavily rely on specific functions for performance tuning or refactoring.
// AI Prompts: Write C# code with Aspose.Cells that creates a "Summary" sheet reporting formula counts grouped by function name. | Modify the example to also include a total formula count per worksheet in the summary. | Explain how to handle nested formulas when extracting function names for counting with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, skips any existing "Summary" tab, scans each worksheet for formula cells, extracts the function name, tallies occurrences per sheet, adds a new "Summary" worksheet with columns for worksheet name, function, and count, auto‑fits columns, and saves the updated file.
class FormulaSummaryExporter
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output_with_summary.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Dictionary to hold counts: Worksheet -> (Function -> Count)
        var summary = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Skip the summary sheet if it already exists
            if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                continue;

            // Prepare inner dictionary for this worksheet
            var funcCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            summary[sheet.Name] = funcCounts;

            // Enumerate all cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                if (!cell.IsFormula) continue; // Only interested in formula cells

                string formula = cell.Formula; // e.g., "=SUM(A1:A10)"
                if (string.IsNullOrEmpty(formula) || formula.Length < 2) continue;

                // Remove leading '=' and trim spaces
                string trimmed = formula.Substring(1).TrimStart();

                // Extract function name (characters before first '(' or space)
                int endIdx = trimmed.IndexOfAny(new char[] { '(', ' ' });
                string funcName = endIdx > 0 ? trimmed.Substring(0, endIdx) : trimmed;

                // Normalize to upper case for consistent grouping
                funcName = funcName.ToUpperInvariant();

                // Update count
                if (funcCounts.ContainsKey(funcName))
                    funcCounts[funcName]++;
                else
                    funcCounts[funcName] = 1;
            }
        }

        // Add (or replace) a worksheet named "Summary"
        Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
        summarySheet.Name = "Summary";

        // Write header
        summarySheet.Cells["A1"].PutValue("Worksheet");
        summarySheet.Cells["B1"].PutValue("Function");
        summarySheet.Cells["C1"].PutValue("Formula Count");

        int rowIndex = 1; // zero‑based index; start after header

        // Populate summary data
        foreach (var wsEntry in summary)
        {
            string wsName = wsEntry.Key;
            foreach (var funcEntry in wsEntry.Value)
            {
                summarySheet.Cells[rowIndex, 0].PutValue(wsName);
                summarySheet.Cells[rowIndex, 1].PutValue(funcEntry.Key);
                summarySheet.Cells[rowIndex, 2].PutValue(funcEntry.Value);
                rowIndex++;
            }
        }

        // Auto‑fit columns for better readability
        summarySheet.AutoFitColumns();

        // Save the workbook with the new summary sheet
        workbook.Save(outputPath);
    }
}

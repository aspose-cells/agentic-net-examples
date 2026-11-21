// Title: Generate a summary worksheet that lists each Excel function and its count per worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a summary tab that shows formula function counts per sheet with Aspose.Cells C# | C# Aspose.Cells example to count Excel functions in each worksheet | extract and tally Excel formula names using Aspose.Cells .NET | create a worksheet summary of function usage and auto‑fit columns in Aspose.Cells | iterate through all worksheets and generate function usage report with Aspose.Cells
// Tags: Aspose.Cells aggregate formula usage per worksheet | C# count Excel functions with Aspose.Cells | extract function name from formula using regex Aspose.Cells | add summary worksheet auto‑fit columns Aspose.Cells | iterate worksheets tally formulas Aspose.Cells

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsFormulaSummary
{
    // The code loads an Excel workbook, adds a new worksheet named "Summary", and writes headers for worksheet name, function, and count. It then loops through every worksheet except the summary sheet, scans the used range, identifies formula cells, extracts the function name with a regular expression, and records the occurrence count in a dictionary. After processing each sheet, the aggregated function counts are written to the summary worksheet, the columns are auto‑fitted, and the workbook is saved with the new summary tab.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Add a new worksheet for the summary and name it "Summary"
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header row in the summary sheet
            int summaryRow = 0;
            summarySheet.Cells[summaryRow, 0].PutValue("Worksheet");
            summarySheet.Cells[summaryRow, 1].PutValue("Function");
            summarySheet.Cells[summaryRow, 2].PutValue("Count");
            summaryRow++;

            // Regular expression to extract the function name from a formula
            // Example: "=SUM(A1:B2)" -> "SUM"
            Regex funcRegex = new Regex(@"^=([A-Za-z0-9_]+)", RegexOptions.Compiled);

            // Iterate through all worksheets except the summary sheet
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet ws = workbook.Worksheets[wsIndex];
                if (ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                    continue; // Skip the summary sheet itself

                // Dictionary to hold function name -> occurrence count for the current worksheet
                Dictionary<string, int> functionCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine the used range of the worksheet
                int maxRow = ws.Cells.MaxDataRow;
                int maxCol = ws.Cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = ws.Cells[row, col];
                        if (cell == null || !cell.IsFormula)
                            continue; // Not a formula cell

                        string formula = cell.Formula; // e.g., "=SUM(A1:B2)"
                        if (string.IsNullOrEmpty(formula))
                            continue;

                        // Extract the function name using the regex
                        Match match = funcRegex.Match(formula);
                        if (match.Success)
                        {
                            string funcName = match.Groups[1].Value;
                            if (functionCounts.ContainsKey(funcName))
                                functionCounts[funcName]++;
                            else
                                functionCounts[funcName] = 1;
                        }
                    }
                }

                // Write the aggregated counts for this worksheet into the summary sheet
                foreach (var kvp in functionCounts)
                {
                    summarySheet.Cells[summaryRow, 0].PutValue(ws.Name);
                    summarySheet.Cells[summaryRow, 1].PutValue(kvp.Key);
                    summarySheet.Cells[summaryRow, 2].PutValue(kvp.Value);
                    summaryRow++;
                }
            }

            // Auto-fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook with the summary sheet (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}

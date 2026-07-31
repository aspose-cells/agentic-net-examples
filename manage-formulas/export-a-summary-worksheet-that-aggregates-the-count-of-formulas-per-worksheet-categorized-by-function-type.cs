// Title: C# – Export a Summary Sheet with Formula Function Counts Using Aspose.Cells
// Description: Loads an existing workbook (or creates a new one), adds a "Summary" worksheet, scans all other sheets for formula cells, extracts each function name, tallies occurrences per function per sheet, writes the data (Worksheet, Function, Count) to the summary sheet, and saves the file.
// Keywords: Aspose.Cells | C# | formula count | function usage | summary worksheet | Excel automation | aggregate formula statistics | count formulas by function | Excel workbook analysis
// Common Searches: Aspose.Cells count formulas per function C# | create summary sheet with formula usage Aspose | how to list Excel functions used in each worksheet | C# code to aggregate formula counts in Excel | generate formula statistics workbook Aspose.Cells
// Developer Intent: Generate a worksheet that reports how many times each formula function appears in every sheet of an Excel file.
// Use Cases: Audit complex workbooks to understand the distribution of calculations. | Identify rarely used functions for performance tuning or refactoring. | Document calculation logic by summarizing function usage per worksheet.
// AI Prompts: Write C# code with Aspose.Cells that adds a summary sheet counting each distinct formula function per worksheet. | Extend the sample to include a total formula count column for each worksheet in the summary. | Explain the function‑name extraction logic and suggest improvements for handling array formulas or nested functions.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook (or creates a new one), adds a "Summary" worksheet, scans all other sheets for formula cells, extracts each function name, tallies occurrences per function per sheet, writes the data (Worksheet, Function, Count) to the summary sheet, and saves the file.
class FormulaSummaryExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Add a new worksheet for the summary.
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Dictionary: Worksheet name -> (Function name -> Count)
            var summary = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets except the summary sheet.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                    continue;

                var funcCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine used range to limit iteration.
                int maxRow = ws.Cells.MaxDataRow;
                int maxCol = ws.Cells.MaxDataColumn;

                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        Cell cell = ws.Cells[r, c];
                        // Use Cell.IsFormula to identify formula cells.
                        if (cell != null && cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                            {
                                // Extract function name (text between '=' and first '(' ).
                                int parenIdx = formula.IndexOf('(');
                                string funcName = parenIdx > 1
                                    ? formula.Substring(1, parenIdx - 1).Trim()
                                    : formula.Substring(1).Trim(); // fallback: whole string after '='

                                if (!funcCounts.ContainsKey(funcName))
                                    funcCounts[funcName] = 0;
                                funcCounts[funcName]++;
                            }
                        }
                    }
                }

                if (funcCounts.Count > 0)
                    summary[ws.Name] = funcCounts;
            }

            // Write headers to the summary worksheet.
            int rowIdx = 0;
            summarySheet.Cells[rowIdx, 0].PutValue("Worksheet");
            summarySheet.Cells[rowIdx, 1].PutValue("Function");
            summarySheet.Cells[rowIdx, 2].PutValue("Count");
            rowIdx++;

            // Populate the summary data.
            foreach (var wsEntry in summary)
            {
                string wsName = wsEntry.Key;
                foreach (var funcEntry in wsEntry.Value)
                {
                    summarySheet.Cells[rowIdx, 0].PutValue(wsName);
                    summarySheet.Cells[rowIdx, 1].PutValue(funcEntry.Key);
                    summarySheet.Cells[rowIdx, 2].PutValue(funcEntry.Value);
                    rowIdx++;
                }
            }

            // Save the workbook with the new summary sheet.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

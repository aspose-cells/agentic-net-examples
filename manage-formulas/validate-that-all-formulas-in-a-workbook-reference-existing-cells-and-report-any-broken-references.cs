// Title: Validate Excel workbook formulas for missing worksheets and out‑of‑range cell references using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells find broken formula references in an Excel file | detect missing worksheet references in Excel formulas with Aspose.Cells .NET | report out of range cell references in Excel formulas using Aspose.Cells C#
// Tags: formula reference validation Aspose.Cells | missing worksheet detection in Excel formulas C# | out‑of‑range cell reference check Aspose.Cells | regex based formula parsing Aspose.Cells | broken formula detection .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet and formula cell, extracts cell and range references with a regular expression, verifies that referenced worksheets exist, parses addresses to row/column indices, checks that they fall within Excel's maximum row and column limits, records missing sheets, invalid addresses, or out‑of‑range cells, and outputs a report of any broken formula references.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold broken reference reports
            List<string> brokenReferences = new List<string>();

            // Regular expression to capture cell/range references (with optional sheet name)
            Regex refRegex = new Regex(
                @"(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?",
                RegexOptions.Compiled);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Skip cells without a formula
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;
                    if (string.IsNullOrEmpty(formula)) continue;

                    // Extract all cell/range references from the formula using regex
                    MatchCollection matches = refRegex.Matches(formula);
                    foreach (Match match in matches)
                    {
                        string referenceRaw = match.Value;

                        // Default to current sheet
                        string targetSheetName = sheet.Name;
                        string address = referenceRaw;

                        // Handle sheet‑qualified references (e.g., 'Sheet2'!A1 or Sheet2!A1)
                        if (referenceRaw.Contains("!"))
                        {
                            string[] parts = referenceRaw.Split('!');
                            targetSheetName = parts[0].Trim('\'');
                            address = parts[1];
                        }

                        // Verify that the target sheet exists
                        Worksheet targetSheet = workbook.Worksheets[targetSheetName];
                        if (targetSheet == null)
                        {
                            brokenReferences.Add(
                                $"Worksheet '{sheet.Name}'!{cell.Name} contains formula with missing sheet '{targetSheetName}'.");
                            continue;
                        }

                        // Parse the address (single cell or range)
                        try
                        {
                            int startRow, startColumn, endRow, endColumn;

                            if (address.Contains(":"))
                            {
                                string[] rangeParts = address.Split(':');
                                CellsHelper.CellNameToIndex(rangeParts[0], out startRow, out startColumn);
                                CellsHelper.CellNameToIndex(rangeParts[1], out endRow, out endColumn);
                            }
                            else
                            {
                                CellsHelper.CellNameToIndex(address, out startRow, out startColumn);
                                endRow = startRow;
                                endColumn = startColumn;
                            }

                            // Validate each cell in the area against Excel limits
                            for (int r = startRow; r <= endRow; r++)
                            {
                                for (int c = startColumn; c <= endColumn; c++)
                                {
                                    if (r < 0 || r > SpreadsheetDefaultMaxRow ||
                                        c < 0 || c > SpreadsheetDefaultMaxColumn)
                                    {
                                        brokenReferences.Add(
                                            $"Worksheet '{sheet.Name}'!{cell.Name} contains formula referencing out‑of‑range cell {targetSheetName}!{CellIndexToName(r, c)}.");
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // If parsing the address fails, it's an invalid reference
                            brokenReferences.Add(
                                $"Worksheet '{sheet.Name}'!{cell.Name} contains formula with invalid reference '{referenceRaw}'.");
                        }
                    }
                }
            }

            // Output the results
            if (brokenReferences.Count == 0)
            {
                Console.WriteLine("No broken formula references were found.");
            }
            else
            {
                Console.WriteLine("Broken formula references detected:");
                foreach (string report in brokenReferences)
                {
                    Console.WriteLine(report);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Constants for maximum rows and columns in Excel (0‑based indices)
    const int SpreadsheetDefaultMaxRow = 1048575;   // 1,048,576 rows
    const int SpreadsheetDefaultMaxColumn = 16383; // 16,384 columns

    // Helper to convert zero‑based row/column indices to A1 style address
    static string CellIndexToName(int row, int column)
    {
        return CellsHelper.CellIndexToName(row, column);
    }
}

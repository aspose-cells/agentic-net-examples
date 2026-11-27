// Title: Identify Excel formulas that reference cells outside the worksheet's used range with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find formulas that reference cells beyond the used range | detect out‑of‑range cell references in Excel formulas using Aspose.Cells .NET | how to validate Excel formula references against MaxDataRow and MaxDataColumn in C# | list formulas pointing to non‑existent cells in a workbook with Aspose.Cells | Aspose.Cells check for invalid cell or range references in formulas
// Tags: formula out‑of‑range detection Aspose.Cells | used range bounds validation .NET | cell reference verification C# | invalid formula range detection Aspose.Cells | Excel formula reference checking Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The C# program loads an Excel workbook with Aspose.Cells, iterates each worksheet, extracts all formula cells, parses each formula for single‑cell and range references using a regular expression, converts those references to zero‑based indices, compares them with the worksheet's MaxDataRow and MaxDataColumn, records any references that exceed these limits, and outputs a list of potential data errors indicating out‑of‑range cell or range references.
class FormulaReferenceChecker
{
    static void Main(string[] args)
    {
        // Path to the workbook to be analyzed (can be passed as first argument)
        string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // List to hold detected issues
            List<string> issues = new List<string>();

            // Regular expression to capture cell references and ranges (e.g., A1 or A1:B10)
            Regex refRegex = new Regex(@"([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?", RegexOptions.Compiled);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Determine the used range of the worksheet
                    int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last used row
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last used column

                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (!cell.IsFormula) continue;

                        string formula = cell.Formula;
                        if (string.IsNullOrEmpty(formula)) continue;

                        // Find all cell references/ranges in the formula
                        foreach (Match match in refRegex.Matches(formula))
                        {
                            // Single cell reference
                            string startRef = match.Groups[1].Value;
                            CellsHelper.CellNameToIndex(startRef, out int startRow, out int startCol);

                            // Check single cell reference
                            if (startRow > maxRow || startCol > maxCol)
                            {
                                string address = cell.Name;
                                string message = $"Worksheet '{sheet.Name}': Formula in cell {address} references out‑of‑range cell {startRef}.";
                                issues.Add(message);
                            }

                            // If a range is present, also validate the end cell
                            if (match.Groups[2].Success)
                            {
                                string endRef = match.Groups[2].Value;
                                CellsHelper.CellNameToIndex(endRef, out int endRow, out int endCol);

                                if (endRow > maxRow || endCol > maxCol)
                                {
                                    string address = cell.Name;
                                    string rangeRef = $"{startRef}:{endRef}";
                                    string message = $"Worksheet '{sheet.Name}': Formula in cell {address} references out‑of‑range range {rangeRef}.";
                                    issues.Add(message);
                                }
                            }
                        }
                    }
                }
                catch (Exception wsEx)
                {
                    Console.WriteLine($"Warning: An error occurred while processing worksheet '{sheet.Name}': {wsEx.Message}");
                }
            }

            // Output the detected issues
            if (issues.Count == 0)
            {
                Console.WriteLine("No formulas reference cells outside the used range.");
            }
            else
            {
                Console.WriteLine("Potential data errors detected:");
                foreach (string issue in issues)
                {
                    Console.WriteLine(issue);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

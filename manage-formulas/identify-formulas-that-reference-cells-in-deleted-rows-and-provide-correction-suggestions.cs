// Title: Find and fix Excel formulas that reference deleted rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# detect formulas that point to removed rows | how to update Excel formula references after deleting rows with Aspose.Cells | list broken cell references in a workbook after row deletion using Aspose.Cells | C# replace invalid row references in Excel formulas programmatically | identify formulas referencing non‑existent rows in .NET Aspose.Cells
// Tags: Aspose.Cells detect invalid formula references | C# adjust Excel formula rows after deletion | Aspose.Cells replace deleted row references | Excel formula range correction with Aspose.Cells | identify broken formulas Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace FormulaReferenceChecker
{
    // The example loads an Excel workbook with Aspose.Cells, scans each worksheet for formulas, parses cell and range references, checks them against a list of deleted rows, and prints suggestions to replace references with the nearest existing rows.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Example: define deleted rows per worksheet (zero‑based index)
            // Key: worksheet name, Value: list of deleted row indices
            var deletedRowsMap = new Dictionary<string, List<int>>
            {
                { "Sheet1", new List<int> { 2, 5 } }, // rows 3 and 6 were deleted
                { "Data", new List<int> { 0, 10 } }   // rows 1 and 11 were deleted
            };

            // Regular expression to capture cell references and ranges (e.g., A3 or B2:C5)
            var refRegex = new Regex(@"([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?", RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the list of deleted rows for this sheet (if any)
                if (!deletedRowsMap.TryGetValue(sheet.Name, out List<int> deletedRows))
                {
                    // No deleted rows for this sheet; skip analysis
                    continue;
                }

                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;
                    if (string.IsNullOrEmpty(formula)) continue;

                    bool referencesDeletedRow = false;
                    var suggestions = new List<string>();

                    // Find all references/ranges in the formula
                    foreach (Match match in refRegex.Matches(formula))
                    {
                        string startRef = match.Groups[1].Value;
                        string endRef = match.Groups[2].Success ? match.Groups[2].Value : null;

                        // Single cell reference
                        if (endRef == null)
                        {
                            int referencedRow = CellReferenceHelper.GetRowIndex(startRef);
                            if (deletedRows.Contains(referencedRow))
                            {
                                referencesDeletedRow = true;
                                int newRow = GetNextExistingRow(referencedRow, deletedRows);
                                string newRef = CellReferenceHelper.ChangeRow(startRef, newRow);
                                suggestions.Add($"Replace {startRef} with {newRef}");
                            }
                        }
                        // Range reference
                        else
                        {
                            int startRow = CellReferenceHelper.GetRowIndex(startRef);
                            int endRow = CellReferenceHelper.GetRowIndex(endRef);

                            bool rangeHasDeletedRow = false;
                            for (int r = startRow; r <= endRow; r++)
                            {
                                if (deletedRows.Contains(r))
                                {
                                    rangeHasDeletedRow = true;
                                    break;
                                }
                            }

                            if (rangeHasDeletedRow)
                            {
                                referencesDeletedRow = true;
                                int newStartRow = GetNextExistingRow(startRow, deletedRows);
                                int newEndRow = GetPreviousExistingRow(endRow, deletedRows);
                                string newStartRef = CellReferenceHelper.ChangeRow(startRef, newStartRow);
                                string newEndRef = CellReferenceHelper.ChangeRow(endRef, newEndRow);
                                suggestions.Add($"Replace {startRef}:{endRef} with {newStartRef}:{newEndRef}");
                            }
                        }
                    }

                    if (referencesDeletedRow)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}");
                        Console.WriteLine($"  Original Formula: {formula}");
                        foreach (string s in suggestions)
                        {
                            Console.WriteLine($"  Suggestion: {s}");
                        }
                    }
                }
            }

            // Optionally, save the workbook after applying corrections (not performed here)
            // workbook.Save("output.xlsx");
        }

        // Returns the next row index that is NOT in the deletedRows list.
        private static int GetNextExistingRow(int startRow, List<int> deletedRows)
        {
            int row = startRow + 1;
            while (deletedRows.Contains(row))
                row++;
            return row;
        }

        // Returns the previous row index that is NOT in the deletedRows list.
        private static int GetPreviousExistingRow(int startRow, List<int> deletedRows)
        {
            int row = startRow - 1;
            while (row >= 0 && deletedRows.Contains(row))
                row--;
            return row >= 0 ? row : 0;
        }
    }

    // Helper class for converting between cell references and zero‑based indices.
    static class CellReferenceHelper
    {
        // Converts a cell reference like "C5" to a zero‑based row index (e.g., 4)
        public static int GetRowIndex(string cellRef)
        {
            // Remove column letters
            string rowPart = Regex.Replace(cellRef, @"[A-Z]+", "");
            if (int.TryParse(rowPart, out int rowNumber))
            {
                return rowNumber - 1; // zero‑based
            }
            throw new ArgumentException($"Invalid cell reference: {cellRef}");
        }

        // Changes the row part of a cell reference to a new zero‑based row index.
        public static string ChangeRow(string cellRef, int newZeroBasedRow)
        {
            // Extract column letters
            string columnPart = Regex.Match(cellRef, @"[A-Z]+").Value;
            int newRowNumber = newZeroBasedRow + 1; // convert back to 1‑based
            return $"{columnPart}{newRowNumber}";
        }
    }
}

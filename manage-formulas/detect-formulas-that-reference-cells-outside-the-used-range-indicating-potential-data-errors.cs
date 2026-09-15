// Title: Identify Excel formulas that reference cells outside the worksheet's used range with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to iterate all worksheets, examine each formula cell, and return a collection of formulas that reference cells beyond the worksheet's MaxDisplayRange. | Enhance the detector to also evaluate multi‑cell range references (e.g., A1:B10) and flag those that extend outside the used area. | Create a reusable method that accepts a Workbook object and returns a list of CellInfo objects containing sheet name, cell address, and formula for out‑of‑range references.
// Common Searches: how to find Excel formulas that point to cells outside the used range using Aspose.Cells C# | C# Aspose.Cells detect formulas referencing non‑existent rows or columns | list formulas with out‑of‑range references in a .xlsx file programmatically | validate worksheet formulas against MaxDisplayRange in .NET
// Tags: Aspose.Cells detect out‑of‑range formula references | C# validate Excel formulas against used range | MaxDisplayRange cell reference verification | regex parsing Excel cell addresses in .NET | scan workbook for invalid formula cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, obtains each sheet's MaxDisplayRange, scans all formula cells, extracts cell and range references with a regular expression, parses them to zero‑based row/column indices, flags any reference that lies outside the used area, and outputs the sheet name, cell address, and formula for each problematic entry.
class FormulaOutsideUsedRangeDetector
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
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

        // List to store information about formulas referencing cells outside the used range
        List<string> problematicFormulas = new List<string>();

        // Regular expression to capture cell references (e.g., A1, $B$2) and range references (e.g., A1:B2)
        Regex cellRefRegex = new Regex(@"\$?[A-Za-z]{1,3}\$?\d+", RegexOptions.Compiled);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                // Determine the used range of the current worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int firstRow = usedRange.FirstRow;
                int firstColumn = usedRange.FirstColumn;
                int lastRow = usedRange.RowCount > 0 ? firstRow + usedRange.RowCount - 1 : firstRow;
                int lastColumn = usedRange.ColumnCount > 0 ? firstColumn + usedRange.ColumnCount - 1 : firstColumn;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;
                    if (string.IsNullOrEmpty(formula)) continue;

                    bool referencesOutside = false;

                    // Find all cell references within the formula
                    foreach (Match match in cellRefRegex.Matches(formula))
                    {
                        string reference = match.Value;

                        // Handle possible range references separated by ':'
                        string[] parts = reference.Split(':');
                        foreach (string part in parts)
                        {
                            // Remove any absolute reference symbols ('$')
                            string cleanPart = part.Replace("$", "");

                            // Try to parse the cleaned part as a cell address
                            if (!TryParseCellReference(cleanPart, out int refRow, out int refCol))
                                continue;

                            // Check if the referenced cell lies outside the used range
                            if (refRow < firstRow || refRow > lastRow ||
                                refCol < firstColumn || refCol > lastColumn)
                            {
                                referencesOutside = true;
                                break;
                            }
                        }

                        if (referencesOutside) break;
                    }

                    if (referencesOutside)
                    {
                        string info = $"Sheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}";
                        problematicFormulas.Add(info);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {ex.Message}");
            }
        }

        // Output the results
        if (problematicFormulas.Count > 0)
        {
            Console.WriteLine("Formulas referencing cells outside the used range:");
            foreach (string entry in problematicFormulas)
            {
                Console.WriteLine(entry);
            }
        }
        else
        {
            Console.WriteLine("No formulas reference cells outside the used range.");
        }
    }

    // Parses an Excel cell reference (e.g., "B12") into zero‑based row and column indices.
    private static bool TryParseCellReference(string cellRef, out int row, out int column)
    {
        row = -1;
        column = -1;

        // Match column letters followed by row numbers
        Match m = Regex.Match(cellRef, @"^([A-Za-z]{1,3})(\d+)$");
        if (!m.Success) return false;

        string colLetters = m.Groups[1].Value.ToUpper();
        string rowNumber = m.Groups[2].Value;

        // Convert column letters to a zero‑based index
        int col = 0;
        foreach (char c in colLetters)
        {
            col = col * 26 + (c - 'A' + 1);
        }
        column = col - 1;

        // Convert row number to zero‑based index
        if (!int.TryParse(rowNumber, out int r)) return false;
        row = r - 1;

        return true;
    }
}

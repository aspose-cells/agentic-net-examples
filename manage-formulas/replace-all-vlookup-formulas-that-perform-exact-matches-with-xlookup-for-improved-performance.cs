// Title: Convert exact‑match VLOOKUP formulas to XLOOKUP in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace VLOOKUP with XLOOKUP for exact matches | programmatically change Excel VLOOKUP formulas to XLOOKUP using Aspose.Cells | how to update VLOOKUP FALSE argument to XLOOKUP in a .NET workbook | iterate cells with Aspose.Cells and convert VLOOKUP to XLOOKUP | convert VLOOKUP formulas to XLOOKUP in C# code
// Tags: Aspose.Cells VLOOKUP to XLOOKUP conversion | C# Excel formula replacement | exact match lookup transformation | XLOOKUP formula generation Aspose.Cells | column offset calculation C#

using System;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, iterates through every worksheet and cell, identifies VLOOKUP formulas that use an exact‑match flag (FALSE or 0), parses the lookup value, table range, and column index, computes the corresponding lookup and return ranges using helper methods for column letters and offsets, builds an XLOOKUP formula with match_mode 0, replaces the original VLOOKUP formula, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook wb = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet ws in wb.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in ws.Cells)
            {
                if (cell.IsFormula && cell.Formula.Contains("VLOOKUP"))
                {
                    string formula = cell.Formula;

                    // Extract the argument list inside the parentheses
                    int start = formula.IndexOf('(');
                    int end = formula.LastIndexOf(')');
                    if (start <= 0 || end <= start) continue;

                    string args = formula.Substring(start + 1, end - start - 1);
                    // Simple split – works for typical VLOOKUP formulas without nested commas
                    string[] parts = args.Split(',');

                    if (parts.Length < 3) continue; // Not a valid VLOOKUP

                    // Determine if the fourth argument (range_lookup) forces an exact match
                    bool isExactMatch = false;
                    if (parts.Length >= 4)
                    {
                        string fourth = parts[3].Trim().Trim('\"');
                        if (string.Equals(fourth, "FALSE", StringComparison.OrdinalIgnoreCase) ||
                            fourth == "0")
                        {
                            isExactMatch = true;
                        }
                    }

                    // If not an exact match, skip replacement
                    if (!isExactMatch) continue;

                    // Build the XLOOKUP formula
                    string lookupValue = parts[0].Trim();
                    string tableArray = parts[1].Trim();
                    string colIndexStr = parts[2].Trim();

                    // Expect tableArray in A1:B10 style
                    string[] rangeParts = tableArray.Split(':');
                    if (rangeParts.Length != 2) continue;

                    string startCell = rangeParts[0];
                    string endCell = rangeParts[1];

                    string startCol = GetColumnLetters(startCell);
                    string rowsPart = GetRowNumbers(startCell, endCell);

                    // Determine the return column based on col_index_num
                    if (!int.TryParse(colIndexStr, out int colIndex) || colIndex <= 0) continue;
                    string returnCol = OffsetColumn(startCol, colIndex - 1);

                    string lookupArray = $"{startCol}{rowsPart}";
                    string returnArray = $"{returnCol}{rowsPart}";

                    // XLOOKUP with exact match (match_mode = 0) and default not_found handling
                    string xlookupFormula = $"=XLOOKUP({lookupValue},{lookupArray},{returnArray},\"#N/A\",0)";

                    // Replace the VLOOKUP formula with XLOOKUP
                    cell.Formula = xlookupFormula;
                }
            }
        }

        // Save the modified workbook
        wb.Save("output.xlsx");
    }

    // Extract column letters from a cell reference (e.g., "B12" -> "B")
    static string GetColumnLetters(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && Char.IsLetter(cellRef[i]))
            i++;
        return cellRef.Substring(0, i);
    }

    // Build the row part for a range (e.g., "A2" and "A10" -> "2:10")
    static string GetRowNumbers(string startCell, string endCell)
    {
        string startRow = GetRowNumber(startCell);
        string endRow = GetRowNumber(endCell);
        return $"{startRow}:{endRow}";
    }

    static string GetRowNumber(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && Char.IsLetter(cellRef[i]))
            i++;
        return cellRef.Substring(i);
    }

    // Offset a column letter by a given number (e.g., "A" + 2 -> "C")
    static string OffsetColumn(string col, int offset)
    {
        int number = ColumnNameToNumber(col);
        number += offset;
        return NumberToColumnName(number);
    }

    static int ColumnNameToNumber(string columnName)
    {
        int sum = 0;
        foreach (char c in columnName.ToUpper())
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
    }

    static string NumberToColumnName(int number)
    {
        string columnName = "";
        while (number > 0)
        {
            int remainder = (number - 1) % 26;
            columnName = (char)(remainder + 'A') + columnName;
            number = (number - 1) / 26;
        }
        return columnName;
    }
}

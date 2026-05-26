using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsVlookupReplacement
{
    class Program
    {
        // Convert a VLOOKUP formula to an INDEX‑MATCH formula.
        // Example: =VLOOKUP(A2,$B$2:$C$10,2,FALSE)
        // becomes: =INDEX($B$2:$C$10,MATCH(A2,$B$2:$B$10,0),2)
        static string ConvertVlookupToIndexMatch(string vlookupFormula)
        {
            // Regex to capture the four arguments of VLOOKUP.
            // Groups: 1‑lookup_value, 2‑table_array, 3‑col_index, 4‑range_lookup
            var pattern = @"=VLOOKUP\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^\)]+)\s*\)";
            var match = Regex.Match(vlookupFormula, pattern, RegexOptions.IgnoreCase);
            if (!match.Success) return vlookupFormula; // fallback – should not happen

            string lookupValue = match.Groups[1].Value.Trim();
            string tableArray = match.Groups[2].Value.Trim();
            string colIndex   = match.Groups[3].Value.Trim();
            // range_lookup (TRUE/FALSE) is ignored because INDEX‑MATCH always performs exact match (0).

            // Build the first‑column range for MATCH.
            string firstColumnRange = GetFirstColumnRange(tableArray);

            // Construct the INDEX‑MATCH formula.
            string indexMatch = $"=INDEX({tableArray},MATCH({lookupValue},{firstColumnRange},0),{colIndex})";
            return indexMatch;
        }

        // Given a range like $B$2:$C$10, return the range that represents only the first column:
        // $B$2:$B$10 (preserving any $ signs).
        static string GetFirstColumnRange(string range)
        {
            // Split start and end addresses.
            var parts = range.Split(':');
            if (parts.Length != 2) return range; // unexpected format – return original.

            string start = parts[0];
            string end   = parts[1];

            // Extract column letters and row numbers.
            string startCol = GetColumnPart(start);
            string startRow = GetRowPart(start);
            string endCol   = GetColumnPart(end);
            string endRow   = GetRowPart(end);

            // Use the start column for both sides.
            string firstColRange = $"{startCol}{startRow}:{startCol}{endRow}";
            return firstColRange;
        }

        // Helper to extract column letters (including $) from an address.
        static string GetColumnPart(string address)
        {
            // Remove any leading $.
            address = address.TrimStart('$');
            // Column letters are the leading letters before the first digit.
            var col = Regex.Match(address, @"^[A-Z]+", RegexOptions.IgnoreCase).Value;
            // Preserve $ if it was present.
            if (address.StartsWith("$")) col = "$" + col;
            // Preserve $ before row if present in original.
            if (address.Contains("$")) col = "$" + col;
            return col;
        }

        // Helper to extract row number (including $) from an address.
        static string GetRowPart(string address)
        {
            // Row is the numeric part after the column letters.
            var row = Regex.Match(address, @"\d+$").Value;
            // Preserve $ if it was present.
            if (address.Contains("$" + row)) row = "$" + row;
            return row;
        }

        static void Main()
        {
            // Load an existing workbook (replace with your actual file path).
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula && cell.Formula.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string originalFormula = cell.Formula;
                            string newFormula = ConvertVlookupToIndexMatch(originalFormula);
                            // Set the new formula (no pre‑calculated value needed).
                            cell.SetFormula(newFormula, new FormulaParseOptions());
                        }
                    }
                }
            }

            // Recalculate all formulas after replacement.
            workbook.CalculateFormula();

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
    }
}
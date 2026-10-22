// Title: Convert all VLOOKUP formulas to XLOOKUP in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace VLOOKUP formulas with XLOOKUP across workbook | programmatically change Excel VLOOKUP to XLOOKUP using .NET API | bulk conversion of legacy VLOOKUP to XLOOKUP with Aspose.Cells | C# code to scan worksheets and update VLOOKUP formulas to XLOOKUP | regex based detection of VLOOKUP formulas in Aspose.Cells
// Tags: VLOOKUP to XLOOKUP conversion Aspose.Cells | bulk formula replacement .NET | Excel formula transformation C# | regex VLOOKUP detection Aspose.Cells | column index to letter utility C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace XLookupReplacement
{
    // The example loads an Excel file with Aspose.Cells, iterates through every worksheet and cell, uses a regular expression to identify VLOOKUP formulas, extracts lookup parameters, computes corresponding XLOOKUP ranges, constructs an XLOOKUP expression, replaces the original formula, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Regex to capture VLOOKUP components
                Regex vlookupRegex = new Regex(
                    @"=VLOOKUP\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^,\)]+)\s*\)",
                    RegexOptions.IgnoreCase);

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula && cell.Formula.Contains("VLOOKUP", StringComparison.OrdinalIgnoreCase))
                        {
                            Match match = vlookupRegex.Match(cell.Formula);
                            if (!match.Success) continue;

                            string lookupValue = match.Groups[1].Value.Trim();
                            string tableArray = match.Groups[2].Value.Trim();
                            string colIndexStr = match.Groups[3].Value.Trim();
                            string rangeLookup = match.Groups[4].Value.Trim();

                            // Parse column index
                            if (!int.TryParse(colIndexStr, out int colIndex) || colIndex < 1)
                                continue; // invalid column index

                            // Determine sheet name and range part
                            string sheetName = sheet.Name; // default
                            string rangePart = tableArray;

                            int exclPos = tableArray.IndexOf('!');
                            if (exclPos >= 0)
                            {
                                sheetName = tableArray.Substring(0, exclPos).Trim('\'');
                                rangePart = tableArray.Substring(exclPos + 1);
                            }

                            // Split range into start and end cells
                            string[] rangeBounds = rangePart.Split(':');
                            if (rangeBounds.Length != 2) continue; // malformed range

                            string startCell = rangeBounds[0];
                            string endCell = rangeBounds[1];

                            // Extract column letters and row numbers
                            string startCol = GetColumnLetter(startCell);
                            string endCol = GetColumnLetter(endCell);
                            int startRow = GetRowNumber(startCell);
                            int endRow = GetRowNumber(endCell);

                            // Build lookup_array (first column of the range)
                            string lookupArray = $"'{sheetName}'!{startCol}{startRow}:{startCol}{endRow}";

                            // Determine return column based on colIndex
                            int startColIndex = ColumnLetterToNumber(startCol);
                            int returnColIndex = startColIndex + colIndex - 1;
                            string returnColLetter = ColumnNumberToLetter(returnColIndex);
                            string returnArray = $"'{sheetName}'!{returnColLetter}{startRow}:{returnColLetter}{endRow}";

                            // Construct XLOOKUP formula
                            string xlookupFormula = $"=XLOOKUP({lookupValue}, {lookupArray}, {returnArray}, \"\", 0)";

                            // Apply new formula
                            cell.Formula = xlookupFormula;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Generic error handling (covers file issues, password protection, etc.)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper: Extract column letters from a cell reference (e.g., $A$1 -> A)
        private static string GetColumnLetter(string cellRef)
        {
            var match = Regex.Match(cellRef, @"([A-Za-z]+)");
            return match.Success ? match.Groups[1].Value.ToUpper() : "";
        }

        // Helper: Extract row number from a cell reference (e.g., $A$12 -> 12)
        private static int GetRowNumber(string cellRef)
        {
            var match = Regex.Match(cellRef, @"(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }

        // Convert column letter to 1‑based index (A=1, B=2, ...)
        private static int ColumnLetterToNumber(string column)
        {
            int sum = 0;
            foreach (char c in column.ToUpper())
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }

        // Convert 1‑based column index to letter (1=A, 2=B, ...)
        private static string ColumnNumberToLetter(int number)
        {
            string column = "";
            while (number > 0)
            {
                int remainder = (number - 1) % 26;
                column = (char)('A' + remainder) + column;
                number = (number - remainder) / 26;
            }
            return column;
        }
    }
}

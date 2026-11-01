// Title: Analyze Excel formulas that reference large worksheets and suggest performance optimizations using Aspose.Cells for .NET
// Common Searches: C# code to find Excel formulas that point to worksheets with more than 100,000 rows using Aspose.Cells | how to identify and report formulas referencing large sheets in Aspose.Cells .NET | performance tips for formulas that reference big data ranges in Excel with Aspose.Cells | detect volatile functions in formulas that reference large worksheets in C# | enable multi‑threaded calculation for workbooks containing formulas on large sheets Aspose.Cells
// Tags: formula reference detection large worksheets Aspose.Cells | performance optimization Excel formulas large data .NET | volatile function identification Aspose.Cells | multi‑threaded calculation setting Aspose.Cells | named range usage to limit sheet references Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace FormulaOptimizationDetector
{
    // This C# example loads a workbook with Aspose.Cells, flags worksheets whose used range exceeds 100 000 rows or 1 000 columns, scans every cell for formulas, extracts referenced sheet names via regex, and reports any formula that points to a large worksheet. For each match it prints suggested optimizations such as using named ranges or tables, replacing volatile functions, switching to SUMIFS/COUNTIFS, enabling multi‑threaded calculation, or pre‑computing results. The tool performs analysis only and does not modify the file.
    class Program
    {
        // Thresholds to consider a worksheet as having a very large data set.
        const int LargeRowCount = 100_000;   // Example: more than 100,000 rows used
        const int LargeColumnCount = 1_000;  // Example: more than 1,000 columns used

        static void Main(string[] args)
        {
            // Path to the workbook to analyze.
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found – \"{inputPath}\"");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Store worksheets that are considered large.
                var largeWorksheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Identify large worksheets based on used range.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // UsedRange gives the area that contains data or formatting.
                    var usedRange = sheet.Cells.MaxDisplayRange;
                    int rowCount = usedRange.RowCount;
                    int columnCount = usedRange.ColumnCount;

                    if (rowCount >= LargeRowCount || columnCount >= LargeColumnCount)
                    {
                        largeWorksheets.Add(sheet.Name);
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" is large (Rows: {rowCount}, Columns: {columnCount}).");
                    }
                }

                // Regular expression to capture sheet names in formulas (e.g., Sheet1!A1 or 'My Sheet'!A1).
                Regex sheetRefRegex = new Regex(@"('([^']+)'|([^'!]+))!", RegexOptions.Compiled);

                // Analyze formulas that reference large worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (!cell.IsFormula) continue;

                        string formula = cell.Formula;
                        if (string.IsNullOrEmpty(formula)) continue;

                        // Find all sheet references within the formula.
                        foreach (Match match in sheetRefRegex.Matches(formula))
                        {
                            string sheetName = match.Groups[2].Success
                                ? match.Groups[2].Value          // Quoted sheet name without quotes
                                : match.Groups[3].Value;         // Unquoted sheet name

                            if (largeWorksheets.Contains(sheetName))
                            {
                                Console.WriteLine($"Formula in {sheet.Name}!{cell.Name} references large worksheet \"{sheetName}\".");
                                Console.WriteLine($"  Formula: {formula}");
                                Console.WriteLine("  Suggested optimizations:");
                                Console.WriteLine("    - Use named ranges or tables to limit the reference scope.");
                                Console.WriteLine("    - Replace volatile functions (e.g., INDIRECT, OFFSET) with static references if possible.");
                                Console.WriteLine("    - Consider using SUMIFS/COUNTIFS instead of array formulas over large ranges.");
                                Console.WriteLine("    - Enable multi‑threaded calculation: workbook.Settings.EnableMultiThreadedCalculation = true;");
                                Console.WriteLine("    - If the formula is only needed for reporting, compute it once and store the result.");
                                Console.WriteLine();
                                // Once reported for this cell, move to the next cell.
                                break;
                            }
                        }
                    }
                }

                // No modifications are saved; this tool is for analysis only.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }

        static string ExtractSheetName(string reference)
        {
            if (string.IsNullOrEmpty(reference))
                return null;

            // If reference contains '!' it separates sheet name from cell/range.
            int exclIndex = reference.IndexOf('!');
            if (exclIndex > 0)
            {
                string sheetPart = reference.Substring(0, exclIndex).Trim();

                // Remove surrounding single quotes if present.
                if (sheetPart.StartsWith("'") && sheetPart.EndsWith("'") && sheetPart.Length > 2)
                {
                    sheetPart = sheetPart.Substring(1, sheetPart.Length - 2);
                }

                return sheetPart;
            }

            // No sheet name present (reference is within the same sheet).
            return null;
        }
    }
}

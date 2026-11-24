// Title: C# program to detect formulas that reference large worksheets and provide optimization recommendations using Aspose.Cells
// Common Searches: how to scan an Excel workbook for formulas that reference sheets larger than 100k rows using Aspose.Cells C# | Aspose.Cells C# detect cross-sheet formula references that may degrade performance | C# find formulas that point to large worksheets and get optimization tips with Aspose.Cells | identify and optimize formulas referencing big data ranges in Excel via Aspose.Cells .NET | using Aspose.Cells to list formulas that reference massive sheets and suggest named range usage
// Tags: cross-sheet formula detection Aspose.Cells | large worksheet formula performance C# | Excel formula optimization using named ranges Aspose.Cells | manual calculation mode recommendation Aspose.Cells | limit whole column references Aspose.Cells

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, flags worksheets exceeding 100,000 rows or 10,000 columns as large, iterates through all formula cells, extracts sheet references via regex, and reports any formulas that point to those large sheets. For each occurrence it prints optimization suggestions such as using named ranges or tables, restricting reference ranges, preferring SUMIFS/COUNTIFS, enabling manual calculation mode, and off‑loading heavy calculations to helper sheets.
class FormulaOptimizer
{
    // Thresholds to consider a worksheet as having a very large data set
    const int LargeRowCount = 100000;   // e.g., >100k rows
    const int LargeColumnCount = 10000; // e.g., >10k columns

    static void Main(string[] args)
    {
        // Path to the workbook to analyze
        string workbookPath = "InputWorkbook.xlsx";

        // Load the workbook (Aspose.Cells handles creation/loading internally)
        Workbook workbook = new Workbook(workbookPath);

        // Build a lookup of worksheets that are considered large
        var largeSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Cells.MaxDataRow + 1 > LargeRowCount ||
                sheet.Cells.MaxDataColumn + 1 > LargeColumnCount)
            {
                largeSheets.Add(sheet.Name);
            }
        }

        // If no large sheets exist, nothing to optimize
        if (largeSheets.Count == 0)
        {
            Console.WriteLine("No worksheets exceed the defined large data thresholds.");
            return;
        }

        // Regex to capture sheet references in formulas (e.g., 'Sheet1'!A1 or Sheet2!B5)
        Regex sheetRefRegex = new Regex(@"(?:(?:'(?<sheet>[^']+)'|(?<sheet>[^'!]+))!)", RegexOptions.Compiled);

        // Iterate through all worksheets and their cells containing formulas
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the range that actually contains data (including formulas)
            var maxRow = sheet.Cells.MaxDataRow;
            var maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Find all sheet references inside the formula
                        var matches = sheetRefRegex.Matches(formula);
                        foreach (Match match in matches)
                        {
                            string referencedSheet = match.Groups["sheet"].Value.Trim();

                            // If the referenced sheet is large, suggest optimization
                            if (largeSheets.Contains(referencedSheet))
                            {
                                Console.WriteLine($"Formula in {sheet.Name}!{cell.Name} references large sheet '{referencedSheet}'.");
                                Console.WriteLine("Optimization suggestions:");
                                Console.WriteLine("- Use named ranges or tables to limit the referenced area.");
                                Console.WriteLine("- Replace whole‑column/whole‑row references with specific ranges.");
                                Console.WriteLine("- Consider using SUMIFS/COUNTIFS with criteria ranges instead of array formulas.");
                                Console.WriteLine("- Enable manual calculation mode and recalculate only when needed.");
                                Console.WriteLine("- If possible, move heavy calculations to a separate helper sheet.");
                                Console.WriteLine();
                                // Once we have reported for this cell, break to avoid duplicate messages
                                break;
                            }
                        }
                    }
                }
            }
        }

        // Optionally, save a copy of the workbook with comments added (not required by the task)
        // workbook.Save("OptimizedWorkbook.xlsx");
    }
}

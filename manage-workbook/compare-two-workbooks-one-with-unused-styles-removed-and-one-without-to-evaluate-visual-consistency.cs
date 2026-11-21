// Title: Compare two Excel workbooks for value and style differences using Aspose.Cells in C# (.NET)
// Common Searches: compare two Excel files cell by cell with Aspose.Cells C# | detect formatting mismatches between workbooks using Aspose.Cells .NET | Aspose.Cells visual consistency check for .xlsx files in C# | C# code to compare worksheet values and styles in Excel workbooks | verify that removing unused styles does not change workbook appearance with Aspose.Cells
// Tags: cell-by-cell workbook comparison Aspose.Cells | style mismatch detection Excel .NET | worksheet range comparison Aspose.Cells | visual consistency verification Excel files C# | unused style impact analysis Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads two .xlsx workbooks, validates their existence, and uses Aspose.Cells for .NET to compare them worksheet by worksheet. It checks worksheet counts and names, determines the maximum used range, then iterates through each cell to compare values and key style attributes such as font, size, color, fill, alignment, and text wrapping. All differences are collected and printed, allowing developers to confirm that removing unused styles does not affect the visual appearance of the workbook.
class WorkbookComparer
{
    static void Main()
    {
        // Paths to the workbooks to compare
        string pathWithUnusedStyles = "Workbook_WithUnusedStyles.xlsx";
        string pathWithoutUnusedStyles = "Workbook_WithoutUnusedStyles.xlsx";

        try
        {
            // Verify that the input files exist
            if (!File.Exists(pathWithUnusedStyles))
                throw new FileNotFoundException($"File not found: {pathWithUnusedStyles}");
            if (!File.Exists(pathWithoutUnusedStyles))
                throw new FileNotFoundException($"File not found: {pathWithoutUnusedStyles}");

            // Load the two workbooks
            Workbook wbWith = new Workbook(pathWithUnusedStyles);
            Workbook wbWithout = new Workbook(pathWithoutUnusedStyles);

            // Perform visual comparison
            List<string> differences = CompareWorkbooks(wbWith, wbWithout);

            // Output the comparison result
            if (differences.Count == 0)
            {
                Console.WriteLine("Workbooks are visually consistent.");
            }
            else
            {
                Console.WriteLine("Differences found:");
                foreach (string diff in differences)
                {
                    Console.WriteLine(diff);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Compares two workbooks cell‑by‑cell for value and style consistency
    static List<string> CompareWorkbooks(Workbook wb1, Workbook wb2)
    {
        var diffs = new List<string>();

        try
        {
            // Compare worksheet counts
            if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            {
                diffs.Add($"Worksheet count mismatch: {wb1.Worksheets.Count} vs {wb2.Worksheets.Count}");
            }

            int sheetCount = Math.Min(wb1.Worksheets.Count, wb2.Worksheets.Count);
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet ws1 = wb1.Worksheets[i];
                Worksheet ws2 = wb2.Worksheets[i];

                // Compare worksheet names
                if (ws1.Name != ws2.Name)
                    diffs.Add($"Worksheet {i} name mismatch: '{ws1.Name}' vs '{ws2.Name}'");

                // Determine the used range for each sheet
                AsposeRange used1 = ws1.Cells.MaxDisplayRange;
                AsposeRange used2 = ws2.Cells.MaxDisplayRange;

                int endRow1 = (used1 != null) ? used1.FirstRow + used1.RowCount - 1 : -1;
                int endRow2 = (used2 != null) ? used2.FirstRow + used2.RowCount - 1 : -1;
                int endCol1 = (used1 != null) ? used1.FirstColumn + used1.ColumnCount - 1 : -1;
                int endCol2 = (used2 != null) ? used2.FirstColumn + used2.ColumnCount - 1 : -1;

                int maxRow = Math.Max(endRow1, endRow2);
                int maxCol = Math.Max(endCol1, endCol2);

                // Iterate through all cells within the combined used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell1 = ws1.Cells[row, col];
                        Cell cell2 = ws2.Cells[row, col];

                        // Compare cell values (including null/empty)
                        object val1 = cell1.Value;
                        object val2 = cell2.Value;
                        if (!Equals(val1, val2))
                        {
                            diffs.Add($"Value mismatch at {ws1.Name}!{CellReference(row, col)}: '{val1}' vs '{val2}'");
                        }

                        // Retrieve styles
                        Style style1 = cell1.GetStyle();
                        Style style2 = cell2.GetStyle();

                        // Compare key style attributes
                        if (!style1.Font.Name.Equals(style2.Font.Name) ||
                            style1.Font.Size != style2.Font.Size ||
                            style1.Font.IsBold != style2.Font.IsBold ||
                            style1.Font.Color != style2.Font.Color ||
                            style1.ForegroundColor != style2.ForegroundColor ||
                            style1.BackgroundColor != style2.BackgroundColor ||
                            style1.HorizontalAlignment != style2.HorizontalAlignment ||
                            style1.VerticalAlignment != style2.VerticalAlignment ||
                            style1.IsTextWrapped != style2.IsTextWrapped)
                        {
                            diffs.Add($"Style mismatch at {ws1.Name}!{CellReference(row, col)}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            diffs.Add($"Comparison error: {ex.Message}");
        }

        return diffs;
    }

    // Helper to convert zero‑based row/column indices to A1 notation
    static string CellReference(int row, int col)
    {
        return CellsHelper.CellIndexToName(row, col);
    }
}

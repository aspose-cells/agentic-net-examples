// Title: Compare two Excel workbooks for visual consistency after removing unused styles using Aspose.Cells in C#
// AI Prompts: Write a C# method that loads two .xlsx files with Aspose.Cells and returns true only if every cell’s value, formula, and full style (font, color, border, alignment, number format) are identical. | Create a console utility that prints the first cell address where the worksheets differ and indicates whether the mismatch is in the cell value or a specific style attribute. | Adjust the comparison routine to ignore style objects that are not referenced in either workbook while still confirming that the visible formatting remains the same.
// Common Searches: Aspose.Cells compare two workbooks visual differences C# | C# check if Excel files have identical formatting using Aspose.Cells | how to detect style mismatches between original and optimized .xlsx in C# | compare cell values and formatting of two Excel workbooks programmatically
// Tags: Aspose.Cells visual workbook comparison C# | Excel cell style equality check Aspose | compare .xlsx formatting differences C# | worksheet consistency validation Aspose.Cells | unused style removal impact assessment

using System;
using System.IO;
using Aspose.Cells;

// The example loads an original workbook and an optimized workbook (with unused styles removed), verifies file existence, then iterates through each worksheet to compare names, cell values (including formulas), and a comprehensive set of style properties such as fonts, colors, patterns, alignments, borders, and number formats. It reports the first discrepancy found and outputs whether the workbooks are visually identical.
class WorkbookComparer
{
    // Entry point
    static void Main()
    {
        const string originalPath = "OriginalWorkbook.xlsx";
        const string optimizedPath = "OptimizedWorkbook.xlsx";

        // Verify that the required files exist
        if (!File.Exists(originalPath))
        {
            Console.WriteLine($"File not found: {originalPath}");
            return;
        }

        if (!File.Exists(optimizedPath))
        {
            Console.WriteLine($"File not found: {optimizedPath}");
            return;
        }

        Workbook wbOriginal = null;
        Workbook wbOptimized = null;

        try
        {
            // Load the two workbooks
            wbOriginal = new Workbook(originalPath);   // Workbook without removing unused styles
            wbOptimized = new Workbook(optimizedPath); // Workbook with unused styles removed
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbooks: {ex.Message}");
            return;
        }

        bool areVisuallyIdentical = false;
        try
        {
            // Perform visual comparison
            areVisuallyIdentical = CompareWorkbooks(wbOriginal, wbOptimized);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during comparison: {ex.Message}");
            return;
        }

        // Output result
        Console.WriteLine(areVisuallyIdentical
            ? "The workbooks are visually identical."
            : "The workbooks differ visually.");
    }

    // Compares two workbooks for visual consistency (values + style properties)
    static bool CompareWorkbooks(Workbook wb1, Workbook wb2)
    {
        // Quick check: same number of worksheets?
        if (wb1.Worksheets.Count != wb2.Worksheets.Count)
        {
            Console.WriteLine("Worksheet count mismatch.");
            return false;
        }

        // Iterate through each worksheet
        for (int i = 0; i < wb1.Worksheets.Count; i++)
        {
            Worksheet ws1 = wb1.Worksheets[i];
            Worksheet ws2 = wb2.Worksheets[i];

            // Compare worksheet names
            if (!ws1.Name.Equals(ws2.Name, StringComparison.Ordinal))
            {
                Console.WriteLine($"Worksheet name mismatch at index {i}: '{ws1.Name}' vs '{ws2.Name}'.");
                return false;
            }

            // Determine the used range for both sheets
            int maxRow = Math.Max(ws1.Cells.MaxDataRow, ws2.Cells.MaxDataRow);
            int maxCol = Math.Max(ws1.Cells.MaxDataColumn, ws2.Cells.MaxDataColumn);

            // Iterate through each cell within the union of used ranges
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell1 = ws1.Cells[row, col];
                    Cell cell2 = ws2.Cells[row, col];

                    // Compare cell values (including formulas)
                    if (!object.Equals(cell1.Value, cell2.Value))
                    {
                        Console.WriteLine($"Value mismatch at {ws1.Name}!{CellIndexToName(row, col)}: '{cell1.Value}' vs '{cell2.Value}'.");
                        return false;
                    }

                    // Compare style objects
                    if (!AreStylesEqual(cell1.GetStyle(), cell2.GetStyle()))
                    {
                        Console.WriteLine($"Style mismatch at {ws1.Name}!{CellIndexToName(row, col)}.");
                        return false;
                    }
                }
            }
        }

        // All checks passed
        return true;
    }

    // Helper to compare two Style objects for visual equality
    static bool AreStylesEqual(Style s1, Style s2)
    {
        // Font properties
        if (s1.Font.Name != s2.Font.Name ||
            s1.Font.Size != s2.Font.Size ||
            s1.Font.IsBold != s2.Font.IsBold ||
            s1.Font.IsItalic != s2.Font.IsItalic ||
            s1.Font.Color != s2.Font.Color ||
            s1.Font.Underline != s2.Font.Underline)
            return false;

        // Background and pattern
        if (s1.ForegroundColor != s2.ForegroundColor ||
            s1.BackgroundColor != s2.BackgroundColor ||
            s1.Pattern != s2.Pattern)
            return false;

        // Alignment
        if (s1.HorizontalAlignment != s2.HorizontalAlignment ||
            s1.VerticalAlignment != s2.VerticalAlignment)
            return false;

        // Borders
        if (!AreBordersEqual(s1, s2))
            return false;

        // Number format
        if (s1.Custom != s2.Custom ||
            s1.Number != s2.Number)
            return false;

        // Text wrapping and shrink to fit
        if (s1.IsTextWrapped != s2.IsTextWrapped ||
            s1.ShrinkToFit != s2.ShrinkToFit)
            return false;

        // All relevant style properties match
        return true;
    }

    // Helper to compare border collections of two styles
    static bool AreBordersEqual(Style s1, Style s2)
    {
        foreach (BorderType side in Enum.GetValues(typeof(BorderType)))
        {
            var b1 = s1.Borders[side];
            var b2 = s2.Borders[side];

            if (b1.LineStyle != b2.LineStyle ||
                b1.Color != b2.Color)
                return false;
        }
        return true;
    }

    // Convert zero‑based row/column indices to Excel cell name (e.g., 0,0 -> A1)
    static string CellIndexToName(int row, int col)
    {
        return CellsHelper.CellIndexToName(row, col);
    }
}

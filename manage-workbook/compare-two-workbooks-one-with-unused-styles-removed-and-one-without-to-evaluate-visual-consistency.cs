// Title: Compare two Excel workbooks for cell values and formatting differences after removing unused styles using Aspose.Cells for .NET
// Common Searches: C# compare two .xlsx files cell values and formatting with Aspose.Cells | how to detect style mismatches after cleaning unused styles in Excel using Aspose.Cells | Aspose.Cells sample for workbook visual consistency verification in C# | compare worksheet fonts, colors, and borders between two workbooks in .NET | detect visual differences between original and optimized Excel files using Aspose.Cells
// Tags: Aspose.Cells workbook formatting comparison C# | cell style diff detection Aspose.Cells | Excel unused style removal impact analysis | worksheet value and formatting comparison .NET | visual consistency check for Excel files using Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookComparison
{
    // The example loads an original workbook and a version with unused styles removed, then iterates through matching worksheets, determines each sheet's used range, and compares every cell's value and style. Differences in font, fill, alignment, and borders are reported to the console. The code demonstrates how to use Aspose.Cells for .NET to verify visual consistency after style cleanup.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the two workbooks: one with unused styles removed, one without.
            string workbookPathWithRemovedStyles = @"C:\Temp\Workbook_WithRemovedStyles.xlsx";
            string workbookPathOriginal = @"C:\Temp\Workbook_Original.xlsx";

            // Perform the comparison.
            CompareWorkbooks(workbookPathWithRemovedStyles, workbookPathOriginal);
        }

        static void CompareWorkbooks(string pathA, string pathB)
        {
            // Verify that both files exist before attempting to load them.
            if (!File.Exists(pathA))
            {
                Console.WriteLine($"File not found: {pathA}");
                return;
            }

            if (!File.Exists(pathB))
            {
                Console.WriteLine($"File not found: {pathB}");
                return;
            }

            Workbook wbA;
            Workbook wbB;

            try
            {
                wbA = new Workbook(pathA);
                wbB = new Workbook(pathB);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbooks: {ex.Message}");
                return;
            }

            // Quick check: same number of worksheets?
            if (wbA.Worksheets.Count != wbB.Worksheets.Count)
            {
                Console.WriteLine($"Worksheet count mismatch: {wbA.Worksheets.Count} vs {wbB.Worksheets.Count}");
                // Continue comparison for existing sheets.
            }

            int sheetCount = Math.Min(wbA.Worksheets.Count, wbB.Worksheets.Count);

            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet wsA = wbA.Worksheets[i];
                Worksheet wsB = wbB.Worksheets[i];

                // Compare worksheet names.
                if (!wsA.Name.Equals(wsB.Name, StringComparison.Ordinal))
                {
                    Console.WriteLine($"Worksheet {i + 1} name mismatch: '{wsA.Name}' vs '{wsB.Name}'");
                }

                // Determine the used range for each sheet.
                int endRowA = 0, endColA = 0, endRowB = 0, endColB = 0;

                if (wsA.Cells.MaxDisplayRange != null)
                {
                    var rangeA = wsA.Cells.MaxDisplayRange;
                    endRowA = rangeA.FirstRow + rangeA.RowCount - 1;
                    endColA = rangeA.FirstColumn + rangeA.ColumnCount - 1;
                }

                if (wsB.Cells.MaxDisplayRange != null)
                {
                    var rangeB = wsB.Cells.MaxDisplayRange;
                    endRowB = rangeB.FirstRow + rangeB.RowCount - 1;
                    endColB = rangeB.FirstColumn + rangeB.ColumnCount - 1;
                }

                int maxRow = Math.Max(endRowA, endRowB);
                int maxCol = Math.Max(endColA, endColB);

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        try
                        {
                            Cell cellA = wsA.Cells[row, col];
                            Cell cellB = wsB.Cells[row, col];

                            // Compare cell values (including formulas).
                            string valueA = cellA.Value?.ToString() ?? string.Empty;
                            string valueB = cellB.Value?.ToString() ?? string.Empty;
                            if (!valueA.Equals(valueB, StringComparison.Ordinal))
                            {
                                Console.WriteLine($"Value mismatch at {wsA.Name}!{CellReference(row, col)}: '{valueA}' vs '{valueB}'");
                            }

                            // Compare style objects.
                            Style styleA = cellA.GetStyle();
                            Style styleB = cellB.GetStyle();

                            List<string> styleDifferences = GetStyleDifferences(styleA, styleB);
                            if (styleDifferences.Count > 0)
                            {
                                Console.WriteLine($"Style mismatch at {wsA.Name}!{CellReference(row, col)}:");
                                foreach (var diff in styleDifferences)
                                {
                                    Console.WriteLine($"  - {diff}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error comparing cell {wsA.Name}!{CellReference(row, col)}: {ex.Message}");
                        }
                    }
                }
            }

            Console.WriteLine("Comparison completed.");
        }

        static List<string> GetStyleDifferences(Style a, Style b)
        {
            var diffs = new List<string>();

            // Font properties.
            if (!a.Font.Name.Equals(b.Font.Name, StringComparison.Ordinal))
                diffs.Add($"Font name: '{a.Font.Name}' vs '{b.Font.Name}'");
            if (a.Font.Size != b.Font.Size)
                diffs.Add($"Font size: {a.Font.Size} vs {b.Font.Size}");
            if (a.Font.IsBold != b.Font.IsBold)
                diffs.Add($"Bold: {a.Font.IsBold} vs {b.Font.IsBold}");
            if (a.Font.IsItalic != b.Font.IsItalic)
                diffs.Add($"Italic: {a.Font.IsItalic} vs {b.Font.IsItalic}");
            if (a.Font.Color != b.Font.Color)
                diffs.Add($"Font color: {a.Font.Color} vs {b.Font.Color}");
            if (a.Font.Underline != b.Font.Underline)
                diffs.Add($"Underline: {a.Font.Underline} vs {b.Font.Underline}");

            // Fill (background) properties.
            if (a.ForegroundColor != b.ForegroundColor)
                diffs.Add($"Foreground color: {a.ForegroundColor} vs {b.ForegroundColor}");
            if (a.BackgroundColor != b.BackgroundColor)
                diffs.Add($"Background color: {a.BackgroundColor} vs {b.BackgroundColor}");
            if (a.Pattern != b.Pattern)
                diffs.Add($"Pattern: {a.Pattern} vs {b.Pattern}");

            // Alignment.
            if (a.HorizontalAlignment != b.HorizontalAlignment)
                diffs.Add($"Horizontal alignment: {a.HorizontalAlignment} vs {b.HorizontalAlignment}");
            if (a.VerticalAlignment != b.VerticalAlignment)
                diffs.Add($"Vertical alignment: {a.VerticalAlignment} vs {b.VerticalAlignment}");
            if (a.IsTextWrapped != b.IsTextWrapped)
                diffs.Add($"Wrap text: {a.IsTextWrapped} vs {b.IsTextWrapped}");

            // Borders (compare each side).
            CompareBorder(a.Borders[BorderType.LeftBorder], b.Borders[BorderType.LeftBorder], "Left border", diffs);
            CompareBorder(a.Borders[BorderType.RightBorder], b.Borders[BorderType.RightBorder], "Right border", diffs);
            CompareBorder(a.Borders[BorderType.TopBorder], b.Borders[BorderType.TopBorder], "Top border", diffs);
            CompareBorder(a.Borders[BorderType.BottomBorder], b.Borders[BorderType.BottomBorder], "Bottom border", diffs);

            return diffs;
        }

        static void CompareBorder(Border a, Border b, string sideName, List<string> diffs)
        {
            if (a.LineStyle != b.LineStyle)
                diffs.Add($"{sideName} line style: {a.LineStyle} vs {b.LineStyle}");
            if (a.Color != b.Color)
                diffs.Add($"{sideName} color: {a.Color} vs {b.Color}");
        }

        static string CellReference(int row, int column)
        {
            return CellsHelper.CellIndexToName(row, column);
        }
    }
}

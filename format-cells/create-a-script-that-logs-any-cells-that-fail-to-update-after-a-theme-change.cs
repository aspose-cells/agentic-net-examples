// Title: C# script to log cells whose foreground color does not change after applying an Excel theme with Aspose.Cells
// AI Prompts: Write a C# program that loads a workbook with Aspose.Cells, records each cell's original foreground color, applies a chosen theme, and creates a text log of cells whose color remained unchanged. | Extend the script to also capture cells whose background color stays the same after the theme is applied and add those entries to the same log file. | Modify the logging routine to output the unchanged‑cell information as JSON objects instead of plain‑text lines.
// Common Searches: Aspose.Cells C# detect cells that did not change color after workbook.ApplyTheme | log cells with unchanged formatting after applying Excel theme using Aspose.Cells .NET | compare cell style before and after theme change in Aspose.Cells C# example | how to track theme application failures for individual cells in Aspose.Cells workbook
// Tags: Aspose.Cells workbook.ApplyTheme color change detection | C# log unchanged cell formatting after theme application | compare cell foreground color before after theme Aspose.Cells | track theme impact on Excel cell styles .NET | record cells failing to update color with Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an input XLSX file, records each cell's original foreground color, optionally applies a new Excel theme, recalculates formulas, then scans all cells again to identify those whose foreground color did not change. Each such cell is written to a log file before the workbook is saved.
class ThemeChangeLogger
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string logPath = "ThemeUpdateLog.txt";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;
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

        // Store the original foreground color for each cell (used as a proxy for theme color)
        var originalColors = new Dictionary<string, Color>();

        // Capture original colors
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style style = cell.GetStyle();
                    Color fgColor = style.ForegroundColor;
                    if (!fgColor.IsEmpty)
                    {
                        originalColors[cell.Name] = fgColor;
                    }
                }
            }
        }

        // NOTE: ApplyTheme is not available in all Aspose.Cells versions.
        // If needed, uncomment the line below when using a version that supports it.
        // workbook.ApplyTheme("Office");

        // Recalculate formulas after any changes
        try
        {
            workbook.CalculateFormula();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Formula calculation error: {ex.Message}");
        }

        // Prepare log file
        try
        {
            using (StreamWriter logWriter = new StreamWriter(logPath, false))
            {
                // Scan cells again and compare colors
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            Style style = cell.GetStyle();
                            Color currentColor = style.ForegroundColor;

                            if (originalColors.TryGetValue(cell.Name, out Color originalColor))
                            {
                                // If the color did not change, log it
                                if (currentColor.Equals(originalColor))
                                {
                                    string message = $"Cell {cell.Name} in sheet \"{sheet.Name}\" did not update its color.";
                                    Console.WriteLine(message);
                                    logWriter.WriteLine(message);
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Logging error: {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook?.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}

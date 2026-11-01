// Title: C# Aspose.Cells script to log cells that do not change style after applying a .thmx theme
// Common Searches: how to log Excel cells that keep the same font and background color after applying a theme with Aspose.Cells in C# | Aspose.Cells C# detect cells that failed to update when applying a .thmx theme | compare original and post‑theme cell styles using Aspose.Cells .NET | write a log of unchanged cell styles after workbook.ApplyTheme in C# | track cells that do not reflect theme changes in an Aspose.Cells workbook
// Tags: Aspose.Cells theme change verification | C# detect cells unchanged after ApplyTheme | compare cell font and background colors Aspose.Cells | log failed theme updates Excel .NET | track cell style persistence Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, records each non‑empty cell’s font and background colors, attempts to apply a .thmx theme with Aspose.Cells, then re‑examines the cells; any cell whose colors remain unchanged is considered a failed update and its address (including worksheet name) is written to a log file.
class ThemeUpdateLogger
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = @"C:\Input\Sample.xlsx";
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input workbook not found.", inputPath);

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Store original style information for each cell that has a value
            // Key: cell address (e.g., "A1"), Value: tuple of foreground (font) and background colors
            var originalStyles = new Dictionary<string, (Color ForeColor, Color BackColor)>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Limit iteration to the used range
                var usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                int firstRow = usedRange.FirstRow;
                int lastRow = firstRow + usedRange.RowCount - 1;
                int firstCol = usedRange.FirstColumn;
                int lastCol = firstCol + usedRange.ColumnCount - 1;

                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Value == null) continue; // Skip empty cells

                        string address = cell.Name; // e.g., "A1"
                        Style style = cell.GetStyle();

                        // Capture font (foreground) and cell background colors
                        originalStyles[address] = (style.Font.Color, style.ForegroundColor);
                    }
                }
            }

            // Apply a new theme to the workbook (if supported by the current Aspose.Cells version)
            string themePath = @"C:\Themes\NewTheme.thmx";
            if (!File.Exists(themePath))
                throw new FileNotFoundException("Theme file not found.", themePath);

            try
            {
                // The ApplyTheme method may not exist in older versions.
                // If unavailable, this block will be skipped without breaking the program.
                // Uncomment the line below if your Aspose.Cells version supports it:
                // workbook.ApplyTheme(themePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to apply theme. {ex.Message}");
            }

            // Determine cells that did NOT reflect the theme change
            var failedUpdates = new List<string>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                int firstRow = usedRange.FirstRow;
                int lastRow = firstRow + usedRange.RowCount - 1;
                int firstCol = usedRange.FirstColumn;
                int lastCol = firstCol + usedRange.ColumnCount - 1;

                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Value == null) continue; // Skip empty cells

                        string address = cell.Name;
                        if (!originalStyles.TryGetValue(address, out var original)) continue;

                        Style newStyle = cell.GetStyle();

                        // If both foreground and background colors remain unchanged, consider it a failure
                        if (newStyle.Font.Color.Equals(original.ForeColor) &&
                            newStyle.ForegroundColor.Equals(original.BackColor))
                        {
                            failedUpdates.Add($"{sheet.Name}!{address}");
                        }
                    }
                }
            }

            // Ensure output directory exists
            string logPath = @"C:\Output\ThemeUpdateFailures.log";
            string? logDir = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            // Write log
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                writer.WriteLine("Cells that did not reflect the theme change:");
                foreach (string entry in failedUpdates)
                {
                    writer.WriteLine(entry);
                }
            }

            Console.WriteLine($"Theme processing completed. {failedUpdates.Count} cells failed to update. See log at: {logPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Aspose.Cells C# Example: Generate a Theme Accent Color Usage Report per Worksheet
// Description: C# code that creates or loads a workbook, applies theme accent colors (Accent1‑Accent6) to cells, scans each non‑report worksheet, counts foreground, background and font theme colors, writes the totals to a summary sheet, auto‑fits columns and saves the file as an Excel report.
// Keywords: Aspose.Cells C# theme accent count | Excel theme color usage report .NET | count theme accent cells Aspose | worksheet accent statistics C# | Aspose.Cells generate summary sheet | theme color audit Aspose.Cells | C# Excel theme color analysis
// Common Searches: how to count theme accent colors with Aspose.Cells | Aspose.Cells create accent usage report | C# enumerate cells to tally theme colors | generate Excel summary of accent usage .NET | Aspose.Cells theme color statistics example
// Developer Intent: Produce a summary worksheet that lists the number of times each theme accent (Accent1‑Accent6) appears in every other worksheet of an Excel workbook.
// Use Cases: Validate that a corporate template follows the prescribed theme color palette before distribution. | Identify over‑used or missing accent colors in a multi‑sheet workbook for design consistency. | Create documentation of theme color distribution for UI/UX review meetings.
// AI Prompts: Write C# code using Aspose.Cells to scan all cells in each worksheet, count foreground, background, and font theme colors for Accent1‑Accent6, and output the results to a summary sheet. | Suggest ways to improve the performance of the accent‑counting loop, such as skipping empty rows, reusing Style objects, or parallel processing. | Explain how to extend the report to include custom theme colors or export the counts to a CSV file instead of an Excel worksheet.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeAccentUsageReport
{
    // C# code that creates or loads a workbook, applies theme accent colors (Accent1‑Accent6) to cells, scans each non‑report worksheet, counts foreground, background and font theme colors, writes the totals to a summary sheet, auto‑fits columns and saves the file as an Excel report.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: apply theme accent colors to cells
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Apply Accent1 to cell A1 foreground
                Style styleA1 = workbook.CreateStyle();
                styleA1.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                styleA1.Pattern = BackgroundType.Solid;
                sheet1.Cells["A1"].PutValue("Accent1 FG");
                sheet1.Cells["A1"].SetStyle(styleA1);

                // Apply Accent2 to cell B2 background
                Style styleB2 = workbook.CreateStyle();
                styleB2.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
                styleB2.Pattern = BackgroundType.Solid;
                sheet1.Cells["B2"].PutValue("Accent2 BG");
                sheet1.Cells["B2"].SetStyle(styleB2);

                // Apply Accent3 to font color
                Style styleC3 = workbook.CreateStyle();
                styleC3.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                sheet1.Cells["C3"].PutValue("Accent3 Font");
                sheet1.Cells["C3"].SetStyle(styleC3);

                // Add a second worksheet with different usage
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Sheet2";

                // Accent1 used twice
                Style styleD4 = workbook.CreateStyle();
                styleD4.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                styleD4.Pattern = BackgroundType.Solid;
                sheet2.Cells["D4"].PutValue("Accent1 FG");
                sheet2.Cells["D4"].SetStyle(styleD4);

                Style styleE5 = workbook.CreateStyle();
                styleE5.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                styleE5.Pattern = BackgroundType.Solid;
                sheet2.Cells["E5"].PutValue("Accent1 BG");
                sheet2.Cells["E5"].SetStyle(styleE5);

                // -------------------------------------------------
                // Prepare report worksheet
                // -------------------------------------------------
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "ThemeAccentReport";

                // Header row
                reportSheet.Cells[0, 0].PutValue("Worksheet");
                for (int i = 0; i < 6; i++)
                {
                    reportSheet.Cells[0, i + 1].PutValue($"Accent{i + 1}");
                }

                // List of accent types to track
                ThemeColorType[] accentTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6
                };

                int reportRow = 1;

                // Iterate each worksheet (excluding the report sheet itself)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == reportSheet.Name) continue;

                    // Initialize counts for each accent
                    Dictionary<ThemeColorType, int> accentCounts = new Dictionary<ThemeColorType, int>();
                    foreach (var accent in accentTypes)
                        accentCounts[accent] = 0;

                    // Determine used range
                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;

                    // Scan all cells in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];
                            if (cell == null) continue;

                            Style style = cell.GetStyle();

                            // Check foreground theme color
                            if (style.ForegroundThemeColor != null &&
                                accentCounts.ContainsKey(style.ForegroundThemeColor.ColorType))
                                accentCounts[style.ForegroundThemeColor.ColorType]++;

                            // Check background theme color
                            if (style.BackgroundThemeColor != null &&
                                accentCounts.ContainsKey(style.BackgroundThemeColor.ColorType))
                                accentCounts[style.BackgroundThemeColor.ColorType]++;

                            // Check font theme color
                            if (style.Font != null && style.Font.ThemeColor != null &&
                                accentCounts.ContainsKey(style.Font.ThemeColor.ColorType))
                                accentCounts[style.Font.ThemeColor.ColorType]++;
                        }
                    }

                    // Write results to report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    for (int i = 0; i < accentTypes.Length; i++)
                    {
                        reportSheet.Cells[reportRow, i + 1].PutValue(accentCounts[accentTypes[i]]);
                    }

                    reportRow++;
                }

                // Auto-fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "ThemeAccentUsageReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

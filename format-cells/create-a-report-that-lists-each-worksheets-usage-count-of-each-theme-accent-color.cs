// Title: Aspose.Cells for .NET – Generate Theme Accent Usage Report per Worksheet (C#)
// Description: C# program using Aspose.Cells that scans each worksheet, counts foreground and background ThemeColorType.Accent1‑Accent6 cells, writes the totals to a "ThemeAccentReport" sheet, auto‑fits columns, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | theme accent colors | ThemeColorType | Excel cell style | count accent usage | worksheet report | auto fit columns | save workbook | Aspose.Cells API
// Common Searches: Aspose.Cells count theme accent colors per sheet | C# generate Excel accent usage report | How to list theme accent usage in workbook using Aspose | Count foreground and background theme colors Aspose.Cells | Create summary of Excel theme colors with Aspose
// Developer Intent: Generate a per‑worksheet report of ThemeColorType accent usage in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Audit workbook branding compliance by summarizing accent color distribution per sheet | Prepare design‑review documentation showing how theme accents are applied across worksheets | Validate template conformity by flagging sheets that exceed allowed accent usage | Automate reporting of theme color statistics for large Excel files
// AI Prompts: Write C# code with Aspose.Cells that iterates all worksheets, counts both foreground and background ThemeColorType.Accent1‑Accent6 cells, and writes the totals to a new 'ThemeAccentReport' sheet. | Adjust the code to count only foreground theme colors and ignore background styles. | Add conditional formatting to the report to highlight rows where any accent count exceeds a specified threshold. | Extend the script to export the accent usage summary to CSV in addition to the Excel report. | Create a reusable method that returns a dictionary of accent counts for a given worksheet.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace ThemeAccentUsageReport
{
    // C# program using Aspose.Cells that scans each worksheet, counts foreground and background ThemeColorType.Accent1‑Accent6 cells, writes the totals to a "ThemeAccentReport" sheet, auto‑fits columns, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load existing template if present; otherwise create a new workbook.
                Workbook workbook;
                string templatePath = "Template.xlsx";

                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // ------------------------------------------------------------
                // Sample data creation – can be removed when using a real workbook.
                // ------------------------------------------------------------
                // Ensure at least one worksheet exists.
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }

                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Apply Accent1 to a few cells.
                Style accent1Style = workbook.CreateStyle();
                accent1Style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                accent1Style.Pattern = BackgroundType.Solid;
                sheet1.Cells["A1"].PutValue("Accent1");
                sheet1.Cells["A1"].SetStyle(accent1Style);
                sheet1.Cells["B2"].PutValue("Accent1");
                sheet1.Cells["B2"].SetStyle(accent1Style);

                // Apply Accent2 to a few cells.
                Style accent2Style = workbook.CreateStyle();
                accent2Style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);
                accent2Style.Pattern = BackgroundType.Solid;
                sheet1.Cells["C3"].PutValue("Accent2");
                sheet1.Cells["C3"].SetStyle(accent2Style);

                // Add a second worksheet with different usage.
                int sheet2Idx = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
                sheet2.Name = "Sheet2";

                // Apply Accent3.
                Style accent3Style = workbook.CreateStyle();
                accent3Style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                accent3Style.Pattern = BackgroundType.Solid;
                sheet2.Cells["A1"].PutValue("Accent3");
                sheet2.Cells["A1"].SetStyle(accent3Style);

                // Apply Accent4.
                Style accent4Style = workbook.CreateStyle();
                accent4Style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0);
                accent4Style.Pattern = BackgroundType.Solid;
                sheet2.Cells["B2"].PutValue("Accent4");
                sheet2.Cells["B2"].SetStyle(accent4Style);
                // ------------------------------------------------------------

                // Theme accent types to analyze.
                ThemeColorType[] accentTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6
                };

                // Create a worksheet to hold the report.
                int reportIdx = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIdx];
                reportSheet.Name = "ThemeAccentReport";
                int reportRow = 0;

                // Header row.
                reportSheet.Cells[reportRow, 0].PutValue("Worksheet");
                for (int i = 0; i < accentTypes.Length; i++)
                {
                    reportSheet.Cells[reportRow, i + 1].PutValue(accentTypes[i].ToString());
                }
                reportRow++;

                // Analyze each worksheet (skip the report sheet itself).
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == reportSheet.Name)
                        continue;

                    // Initialise counters for each accent.
                    var accentCounts = new Dictionary<ThemeColorType, int>();
                    foreach (ThemeColorType t in accentTypes)
                    {
                        accentCounts[t] = 0;
                    }

                    // Determine used range.
                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;

                    // Scan cells within the used range.
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];
                            Style style = cell.GetStyle();

                            // Foreground theme color.
                            if (style.ForegroundThemeColor != null &&
                                accentCounts.ContainsKey(style.ForegroundThemeColor.ColorType))
                            {
                                accentCounts[style.ForegroundThemeColor.ColorType]++;
                            }

                            // Background theme color.
                            if (style.BackgroundThemeColor != null &&
                                accentCounts.ContainsKey(style.BackgroundThemeColor.ColorType))
                            {
                                accentCounts[style.BackgroundThemeColor.ColorType]++;
                            }
                        }
                    }

                    // Write results for this worksheet.
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    for (int i = 0; i < accentTypes.Length; i++)
                    {
                        reportSheet.Cells[reportRow, i + 1].PutValue(accentCounts[accentTypes[i]]);
                    }
                    reportRow++;
                }

                // Auto‑fit columns for readability.
                reportSheet.AutoFitColumns();

                // Save the workbook with the report.
                string outputPath = "ThemeAccentUsageReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the theme accent usage report:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

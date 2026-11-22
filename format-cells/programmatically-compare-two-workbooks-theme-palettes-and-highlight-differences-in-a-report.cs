// Title: Generate a theme palette comparison report for two Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells compare theme colors between two .xlsx files | how to create a theme palette difference report with Aspose.Cells | Aspose.Cells GetThemeColor example for workbook comparison | highlight mismatched theme colors in Excel using Aspose.Cells | save theme palette comparison results to a new workbook in C#
// Tags: Aspose.Cells GetThemeColor API usage | compare Excel workbook theme palettes | generate Excel report of theme color differences | apply conditional row styling Aspose.Cells | auto‑fit columns and save workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ThemePaletteComparison
{
    // The sample loads two Excel workbooks, iterates through all ThemeColorType values, retrieves each workbook's theme colors via Workbook.GetThemeColor, and records any mismatches. It then creates a new workbook, writes a header row, lists each differing theme color with its ARGB values, highlights those rows with a light salmon background, adds a note when no differences are found, auto‑fits the columns, ensures the output directory exists, and saves the comparison report to the specified path.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source workbooks and the output report
            string workbookPath1 = @"C:\Data\Workbook1.xlsx";
            string workbookPath2 = @"C:\Data\Workbook2.xlsx";
            string reportPath = @"C:\Data\ThemePaletteReport.xlsx";

            try
            {
                // Verify that source files exist
                if (!File.Exists(workbookPath1))
                    throw new FileNotFoundException($"Source workbook not found: {workbookPath1}");
                if (!File.Exists(workbookPath2))
                    throw new FileNotFoundException($"Source workbook not found: {workbookPath2}");

                // Load the two workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // All theme color types that are part of a palette
                ThemeColorType[] themeColors = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6,
                    ThemeColorType.Background1,
                    ThemeColorType.Background2,
                    ThemeColorType.Text1,
                    ThemeColorType.Text2,
                    ThemeColorType.Hyperlink,
                    ThemeColorType.FollowedHyperlink
                };

                // List to hold differences
                List<(ThemeColorType Type, Color Color1, Color Color2)> differences = new List<(ThemeColorType, Color, Color)>();

                // Compare each theme color using Workbook.GetThemeColor
                foreach (ThemeColorType type in themeColors)
                {
                    Color color1 = wb1.GetThemeColor(type);
                    Color color2 = wb2.GetThemeColor(type);

                    if (!color1.Equals(color2))
                    {
                        differences.Add((type, color1, color2));
                    }
                }

                // Create a new workbook for the report
                Workbook reportWb = new Workbook();
                Worksheet sheet = reportWb.Worksheets[0];
                sheet.Name = "Theme Palette Differences";

                // Write header
                sheet.Cells[0, 0].PutValue("Theme Color");
                sheet.Cells[0, 1].PutValue("Workbook1 Color (ARGB)");
                sheet.Cells[0, 2].PutValue("Workbook2 Color (ARGB)");
                sheet.Cells[0, 3].PutValue("Difference");

                // Apply bold style to header
                Style headerStyle = reportWb.CreateStyle();
                headerStyle.Font.IsBold = true;
                StyleFlag flag = new StyleFlag { All = true };
                sheet.Cells.CreateRange(0, 0, 1, 4).ApplyStyle(headerStyle, flag);

                // Populate differences
                int rowIndex = 1;
                foreach (var diff in differences)
                {
                    sheet.Cells[rowIndex, 0].PutValue(diff.Type.ToString());
                    sheet.Cells[rowIndex, 1].PutValue(diff.Color1.ToArgb());
                    sheet.Cells[rowIndex, 2].PutValue(diff.Color2.ToArgb());
                    sheet.Cells[rowIndex, 3].PutValue("Different");

                    // Highlight the row with a light red background
                    Style diffStyle = reportWb.CreateStyle();
                    diffStyle.ForegroundColor = Color.LightSalmon;
                    diffStyle.Pattern = BackgroundType.Solid;
                    sheet.Cells.CreateRange(rowIndex, 0, 1, 4).ApplyStyle(diffStyle, flag);

                    rowIndex++;
                }

                // If no differences were found, write a note
                if (differences.Count == 0)
                {
                    sheet.Cells[rowIndex, 0].PutValue("No differences found between the two theme palettes.");
                    sheet.Cells.Merge(rowIndex, 0, 1, 4);
                    Style noteStyle = reportWb.CreateStyle();
                    noteStyle.Font.IsBold = true;
                    sheet.Cells[rowIndex, 0].SetStyle(noteStyle);
                }

                // Auto-fit columns for better readability
                sheet.AutoFitColumns();

                // Ensure the output directory exists
                string reportDir = Path.GetDirectoryName(reportPath);
                if (!Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                // Save the report workbook
                try
                {
                    reportWb.Save(reportPath);
                    Console.WriteLine($"Report saved to: {reportPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Failed to save the report:");
                    Console.WriteLine(saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

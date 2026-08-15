// Title: Compare Excel Theme Palettes and Create a Highlighted Report with Aspose.Cells (C#)
// Description: Loads two Excel workbooks, extracts the first 12 ThemeColorType values using Aspose.Cells, writes each color type and its hex code to a new worksheet, marks rows with mismatched colors in light‑salmon, auto‑fits columns, and saves the comparison file.
// Keywords: Aspose.Cells theme palette comparison | C# GetThemeColor example | Excel theme color differences | highlight mismatched theme colors | generate theme comparison report | compare workbook themes Aspose | Excel branding validation C#
// Common Searches: compare theme colors of two Excel files Aspose.Cells | C# generate theme palette report | highlight rows when Excel theme colors differ | Aspose.Cells GetThemeColor usage | Excel theme palette validation script
// Developer Intent: Produce a new workbook that lists each theme color type from two source files, shows their hex values side‑by‑side, and highlights rows where the colors are not identical.
// Use Cases: Verify that multiple Excel templates follow a corporate color scheme by comparing their theme palettes. | Automate a quality‑check in a document‑generation pipeline to ensure generated files use the same theme as a master workbook. | Provide designers with a side‑by‑side Excel report that flags any color deviations between versioned spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells that reads two workbooks, compares their theme palettes, and creates a formatted report highlighting differences. | Generate a helper method that returns a dictionary of ThemeColorType to hex string for a given Workbook using Aspose.Cells. | Explain how to apply conditional row styling in Aspose.Cells based on a boolean comparison of theme colors.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemePaletteComparison
{
    // Loads two Excel workbooks, extracts the first 12 ThemeColorType values using Aspose.Cells, writes each color type and its hex code to a new worksheet, marks rows with mismatched colors in light‑salmon, auto‑fits columns, and saves the comparison file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to the workbooks to compare
                string workbookPath1 = "Workbook1.xlsx";
                string workbookPath2 = "Workbook2.xlsx";

                // Verify that the input files exist
                if (!File.Exists(workbookPath1))
                {
                    Console.WriteLine($"Error: File not found - {workbookPath1}");
                    return;
                }
                if (!File.Exists(workbookPath2))
                {
                    Console.WriteLine($"Error: File not found - {workbookPath2}");
                    return;
                }

                // Load the two workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // Create a new workbook that will hold the comparison report
                Workbook report = new Workbook();
                Worksheet sheet = report.Worksheets[0];
                sheet.Name = "Theme Comparison";

                // Write header row
                sheet.Cells["A1"].PutValue("Theme Color Type");
                sheet.Cells["B1"].PutValue("Workbook 1 Color");
                sheet.Cells["C1"].PutValue("Workbook 2 Color");
                sheet.Cells["D1"].PutValue("Difference");

                // Apply bold style to header
                Style headerStyle = report.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Size = 12;
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.ForegroundColor = Color.LightGray;
                StyleFlag headerFlag = new StyleFlag { All = true };
                sheet.Cells["A1:D1"].SetStyle(headerStyle, headerFlag);

                int rowIndex = 1; // zero‑based index; row 2 in Excel

                // Iterate through all theme color types (0‑11)
                foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
                {
                    // Only process the first 12 defined types (skip StyleColor)
                    if ((int)type > 11) break;

                    // Retrieve theme colors from both workbooks
                    Color color1 = wb1.GetThemeColor(type);
                    Color color2 = wb2.GetThemeColor(type);

                    // Write data to the report sheet
                    sheet.Cells[rowIndex, 0].PutValue(type.ToString());
                    sheet.Cells[rowIndex, 1].PutValue(ColorToHex(color1));
                    sheet.Cells[rowIndex, 2].PutValue(ColorToHex(color2));

                    // Determine if colors differ
                    bool isDifferent = !color1.Equals(color2);
                    sheet.Cells[rowIndex, 3].PutValue(isDifferent ? "Yes" : "No");

                    // Highlight the entire row if there is a difference
                    if (isDifferent)
                    {
                        Style diffStyle = report.CreateStyle();
                        diffStyle.Pattern = BackgroundType.Solid;
                        diffStyle.ForegroundColor = Color.LightSalmon;
                        sheet.Cells.CreateRange(rowIndex, 0, 1, 4).SetStyle(diffStyle);
                    }

                    rowIndex++;
                }

                // Auto‑fit columns for better readability
                sheet.AutoFitColumns();

                // Save the comparison report
                string reportPath = "ThemePaletteComparisonReport.xlsx";
                report.Save(reportPath);
                Console.WriteLine($"Report saved to {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert a Color to a hex string (e.g., #RRGGBB)
        private static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}

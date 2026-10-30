// Title: Create a C# program with Aspose.Cells to compare the theme palettes of two Excel workbooks and generate a highlighted difference report
// Common Searches: C# Aspose.Cells compare theme colors of two .xlsx files and output a report | How to use GetThemeColor in Aspose.Cells to find differences between workbook themes | Generate an Excel theme palette comparison worksheet with Aspose.Cells .NET | Highlight mismatched theme colors in a generated report using Aspose.Cells C# | Auto-fit columns after creating a theme comparison report with Aspose.Cells
// Tags: compare workbook theme palettes Aspose.Cells | Excel theme palette difference report C# | theme color mismatch highlighting Aspose.Cells | create theme comparison worksheet Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads two Excel workbooks, extracts each of the 12 theme colors using GetThemeColor, writes the index and ARGB values to a new worksheet, applies the actual colors as cell backgrounds, marks cells where the colors differ, auto‑fits the columns, and saves the file as a theme‑comparison report.
class ThemeComparison
{
    static void Main()
    {
        try
        {
            // Verify that the source workbooks exist
            const string file1 = "Workbook1.xlsx";
            const string file2 = "Workbook2.xlsx";

            if (!File.Exists(file1) || !File.Exists(file2))
            {
                Console.WriteLine("One or both source workbook files are missing.");
                return;
            }

            // Load the two workbooks to compare
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Create a new workbook that will hold the comparison report
            Workbook report = new Workbook();
            Worksheet ws = report.Worksheets[0];
            ws.Name = "Theme Comparison";

            // Write header row
            ws.Cells["A1"].PutValue("Color Index");
            ws.Cells["B1"].PutValue("Workbook1 Color");
            ws.Cells["C1"].PutValue("Workbook2 Color");
            ws.Cells["D1"].PutValue("Are Equal");

            // Theme palettes contain 12 colors (indices 0‑11)
            for (int i = 0; i < 12; i++)
            {
                int row = i + 2; // Data starts at row 2

                // Write the color index
                ws.Cells[row, 0].PutValue(i);

                // Retrieve theme colors from both workbooks using GetThemeColor
                ThemeColorType colorType = (ThemeColorType)i;
                Color color1 = wb1.GetThemeColor(colorType);
                Color color2 = wb2.GetThemeColor(colorType);

                // Write ARGB values for reference
                ws.Cells[row, 1].PutValue(color1.ToArgb());
                ws.Cells[row, 2].PutValue(color2.ToArgb());

                // Apply background fill so the cell shows the actual color
                Style style1 = ws.Cells[row, 1].GetStyle();
                style1.ForegroundColor = color1;
                style1.Pattern = BackgroundType.Solid;
                ws.Cells[row, 1].SetStyle(style1);

                Style style2 = ws.Cells[row, 2].GetStyle();
                style2.ForegroundColor = color2;
                style2.Pattern = BackgroundType.Solid;
                ws.Cells[row, 2].SetStyle(style2);

                // Determine if the colors are identical
                bool areEqual = color1.ToArgb() == color2.ToArgb();
                ws.Cells[row, 3].PutValue(areEqual ? "Yes" : "No");

                // Highlight the "Are Equal" cell when colors differ
                if (!areEqual)
                {
                    Style diffStyle = ws.Cells[row, 3].GetStyle();
                    diffStyle.ForegroundColor = Color.Yellow;
                    diffStyle.Pattern = BackgroundType.Solid;
                    ws.Cells[row, 3].SetStyle(diffStyle);
                }
            }

            // Adjust column widths for readability
            ws.AutoFitColumns();

            // Save the comparison report
            const string reportFile = "ThemeComparisonReport.xlsx";
            report.Save(reportFile);
            Console.WriteLine($"Theme comparison report saved to '{reportFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

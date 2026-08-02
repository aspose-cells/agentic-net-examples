using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemePaletteComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source workbooks and the output report
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";
            string reportPath = "ThemePaletteComparisonReport.xlsx";

            CompareThemePalettes(workbookPath1, workbookPath2, reportPath);
        }

        /// <summary>
        /// Compares the theme palettes of two workbooks and generates a report highlighting differences.
        /// </summary>
        /// <param name="path1">Path to the first workbook.</param>
        /// <param name="path2">Path to the second workbook.</param>
        /// <param name="reportPath">Path where the comparison report will be saved.</param>
        static void CompareThemePalettes(string path1, string path2, string reportPath)
        {
            // Load the two workbooks (lifecycle rule: load)
            Workbook wb1 = new Workbook(path1);
            Workbook wb2 = new Workbook(path2);

            // Create a new workbook for the report (lifecycle rule: create)
            Workbook reportWb = new Workbook();
            Worksheet sheet = reportWb.Worksheets[0];

            // Write header row
            sheet.Cells["A1"].PutValue("Theme Color Type");
            sheet.Cells["B1"].PutValue("Workbook 1 Color (ARGB)");
            sheet.Cells["C1"].PutValue("Workbook 2 Color (ARGB)");
            sheet.Cells["D1"].PutValue("Difference");

            // Apply bold style to header
            Style headerStyle = reportWb.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            for (int col = 0; col < 4; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            int rowIndex = 1; // start after header

            // Iterate through all theme color types (0 to 11)
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve theme colors from both workbooks
                Color color1 = wb1.GetThemeColor(type);
                Color color2 = wb2.GetThemeColor(type);

                // Compare colors
                bool isEqual = color1.ToArgb() == color2.ToArgb();

                // Write data only if there is a difference
                if (!isEqual)
                {
                    // Theme color type name
                    sheet.Cells[rowIndex, 0].PutValue(type.ToString());

                    // Color values as ARGB strings
                    sheet.Cells[rowIndex, 1].PutValue($"A={color1.A}, R={color1.R}, G={color1.G}, B={color1.B}");
                    sheet.Cells[rowIndex, 2].PutValue($"A={color2.A}, R={color2.R}, G={color2.G}, B={color2.B}");

                    // Mark difference
                    sheet.Cells[rowIndex, 3].PutValue("Different");

                    // Highlight the cells with the actual colors for visual reference
                    Style style1 = reportWb.CreateStyle();
                    style1.ForegroundColor = color1;
                    style1.Pattern = BackgroundType.Solid;
                    sheet.Cells[rowIndex, 1].SetStyle(style1);

                    Style style2 = reportWb.CreateStyle();
                    style2.ForegroundColor = color2;
                    style2.Pattern = BackgroundType.Solid;
                    sheet.Cells[rowIndex, 2].SetStyle(style2);

                    rowIndex++;
                }
            }

            // If no differences were found, indicate that in the report
            if (rowIndex == 1)
            {
                sheet.Cells[rowIndex, 0].PutValue("No differences found between the theme palettes.");
                // Merge cells A2:D2 for better presentation
                sheet.Cells.CreateRange("A2:D2").Merge();
                Style infoStyle = reportWb.CreateStyle();
                infoStyle.Font.IsBold = true;
                infoStyle.HorizontalAlignment = TextAlignmentType.Center;
                sheet.Cells[rowIndex, 0].SetStyle(infoStyle);
            }

            // Auto-fit columns for readability
            sheet.AutoFitColumns();

            // Save the report workbook (lifecycle rule: save)
            reportWb.Save(reportPath);
        }
    }
}
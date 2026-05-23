using System;
using System.Drawing;
using Aspose.Cells;

class ThemeComparison
{
    static void Main()
    {
        // Paths to the two workbooks to compare
        string wbPath1 = "Workbook1.xlsx";
        string wbPath2 = "Workbook2.xlsx";

        // Load the source workbooks (lifecycle rule: use constructor for loading)
        Workbook wb1 = new Workbook(wbPath1);
        Workbook wb2 = new Workbook(wbPath2);

        // Create a new workbook that will hold the comparison report (lifecycle rule: use default constructor)
        Workbook report = new Workbook();
        Worksheet sheet = report.Worksheets[0];
        sheet.Name = "Theme Comparison";

        // Write header cells
        sheet.Cells["A1"].PutValue("Theme Color Type");
        sheet.Cells["B1"].PutValue("Workbook1 Color");
        sheet.Cells["C1"].PutValue("Workbook2 Color");
        sheet.Cells["D1"].PutValue("Different?");

        // Apply a simple header style (bold, centered)
        Style headerStyle = report.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Size = 12;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        for (int col = 0; col < 4; col++)
        {
            sheet.Cells[0, col].SetStyle(headerStyle);
        }

        int row = 1; // start after header

        // Iterate over the 12 theme color types (Background1 … FollowedHyperlink)
        foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
        {
            // ThemeColorType enum contains extra values; limit to the first 12 (0‑11)
            if ((int)type > 11) break;

            // Retrieve theme colors from both workbooks
            Color color1 = wb1.GetThemeColor(type);
            Color color2 = wb2.GetThemeColor(type);

            // Write the type name
            sheet.Cells[row, 0].PutValue(type.ToString());

            // Write colors as ARGB hex strings for readability
            sheet.Cells[row, 1].PutValue(ColorToString(color1));
            sheet.Cells[row, 2].PutValue(ColorToString(color2));

            // Determine if the colors differ
            bool different = !color1.Equals(color2);
            sheet.Cells[row, 3].PutValue(different ? "Yes" : "No");

            // Highlight the entire row when a difference is found
            if (different)
            {
                Style diffStyle = report.CreateStyle();
                diffStyle.ForegroundColor = Color.LightSalmon;
                diffStyle.Pattern = BackgroundType.Solid;
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].SetStyle(diffStyle);
                }
            }

            row++;
        }

        // Adjust column widths to fit content
        sheet.AutoFitColumns();

        // Save the report workbook (lifecycle rule: use Save method)
        report.Save("ThemeComparisonReport.xlsx");
    }

    // Helper method to convert a Color to a hex ARGB string (e.g., #FF112233)
    static string ColorToString(Color c)
    {
        return $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
    }
}
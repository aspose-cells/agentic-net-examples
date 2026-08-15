// Title: C# – List All Aspose.Cells ThemeColorType Names with Hex Values in Excel
// Description: Creates a new workbook, adds a "Theme Summary" worksheet, writes headers, iterates through every ThemeColorType (except StyleColor), extracts each theme color with Workbook.GetThemeColor, converts it to a #RRGGBB string, records the name and hex code, auto‑fits columns, and saves the file as ThemeSummary.xlsx.
// Keywords: Aspose.Cells ThemeColorType list | C# export theme palette | GetThemeColor hex code | .NET Excel theme colors | generate theme summary worksheet
// Common Searches: list theme colors Aspose.Cells C# | export theme palette to Excel using Aspose | how to get hex values of ThemeColorType | skip StyleColor when extracting theme colors | Aspose.Cells GetThemeColor example
// Developer Intent: Produce an Excel worksheet that enumerates each ThemeColorType name alongside its hexadecimal representation.
// Use Cases: Provide designers a quick reference of the workbook's theme palette. | Include the theme palette in brand style guides or documentation. | Validate workbook colors against corporate branding standards.
// AI Prompts: Generate C# code with Aspose.Cells that writes all ThemeColorType names and #RRGGBB values to a worksheet, excluding StyleColor. | Modify the example to output ARGB hex strings (#AARRGGBB) for each theme color. | Add a column that displays the numeric RGB components (e.g., 255,0,170) next to the hex code.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSummary
{
    // Creates a new workbook, adds a "Theme Summary" worksheet, writes headers, iterates through every ThemeColorType (except StyleColor), extracts each theme color with Workbook.GetThemeColor, converts it to a #RRGGBB string, records the name and hex code, auto‑fits columns, and saves the file as ThemeSummary.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet for the summary
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Theme Summary";

            // Write headers
            summarySheet.Cells["A1"].PutValue("Theme Color Name");
            summarySheet.Cells["B1"].PutValue("Hexadecimal Value");

            // Get all theme color types (Background1 to FollowedHyperlink)
            ThemeColorType[] themeTypes = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));

            int rowIndex = 1; // zero‑based index; row 1 is the second row (A2, B2)

            foreach (ThemeColorType type in themeTypes)
            {
                // Skip the StyleColor entry (value 12) as it is not a real theme color
                if (type == ThemeColorType.StyleColor) continue;

                // Retrieve the theme color from the workbook
                Color color = workbook.GetThemeColor(type);

                // Convert the color to a hex string (e.g., #FF00AA)
                string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

                // Write the name and hex value to the sheet
                summarySheet.Cells[rowIndex, 0].PutValue(type.ToString()); // Column A
                summarySheet.Cells[rowIndex, 1].PutValue(hex);            // Column B

                rowIndex++;
            }

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("ThemeSummary.xlsx");
        }
    }
}

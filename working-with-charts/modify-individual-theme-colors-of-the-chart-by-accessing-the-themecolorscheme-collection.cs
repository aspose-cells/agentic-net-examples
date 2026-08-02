// Title: Change individual chart theme colors with Workbook.SetThemeColor in Aspose.Cells for .NET
// Description: Demonstrates how to modify the workbook's Theme.ColorScheme (Accent1, Accent2, Text1, Hyperlink, etc.) using Workbook.SetThemeColor so that a column chart automatically adopts the new colors, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | Workbook.SetThemeColor | ThemeColorType | chart theme colors | modify Excel theme programmatically | column chart color palette | brand colors in Excel | update workbook theme scheme
// Common Searches: Aspose.Cells change chart theme colors | SetThemeColor example C# | how to update Accent1 Accent2 in Excel with Aspose | modify workbook theme color scheme programmatically | chart series inherit workbook theme Aspose.Cells
// Developer Intent: The developer wants to programmatically adjust specific theme colors (e.g., Accent1, Accent2) so that chart series reflect a custom color palette without setting each series color individually.
// Use Cases: Apply a corporate brand palette by redefining Accent1 and Accent2 before generating charts. | Create themed reports where chart colors, text, and hyperlinks follow a unified style defined in the workbook theme. | Generate multiple workbooks with consistent visual branding by updating the theme once rather than styling each chart element.
// AI Prompts: Show how to change Accent3 and Accent4 theme colors for a chart using Aspose.Cells. | Provide code to reset all workbook theme colors to their default values after modification. | Explain whether existing charts need a refresh after calling Workbook.SetThemeColor and how the change propagates.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeDemo
{
    // Demonstrates how to modify the workbook's Theme.ColorScheme (Accent1, Accent2, Text1, Hyperlink, etc.) using Workbook.SetThemeColor so that a column chart automatically adopts the new colors, then saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 15);
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.NSeries.Add("B1:C6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Initially use theme colors for the series (default Accent1 and Accent2)
            // No explicit color assignment needed; they inherit from the workbook theme.

            // Modify individual theme colors that the chart will use
            // Change Accent1 to a deep orange and Accent2 to a teal color
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(255, 140, 0)); // Deep orange
            workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(0, 128, 128)); // Teal

            // Optionally modify other theme entries, e.g., Text1 and Hyperlink
            workbook.SetThemeColor(ThemeColorType.Text1, Color.DarkSlateGray);
            workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.MediumPurple);

            // Save the workbook to verify that the chart reflects the new theme colors
            workbook.Save("ChartWithModifiedThemeColors.xlsx");
        }
    }
}

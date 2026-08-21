// Title: Aspose.Cells .NET – Change Individual Chart Theme Colors with SetThemeColor
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and customize specific theme colors (Accent1, Accent2, Text1, Background1) using the Workbook.SetThemeColor method, which updates the chart's Theme.ColorScheme before saving the file.
// Keywords: Aspose.Cells set theme color | modify chart theme colors .NET | Workbook.SetThemeColor example | custom Excel theme palette Aspose | change Accent1 Accent2 colors programmatically | Theme.ColorScheme Aspose.Cells | C# chart color customization | Excel theme color API
// Common Searches: how to change chart theme colors with Aspose.Cells | Aspose.Cells SetThemeColor C# example | customize Excel theme palette programmatically | change Accent1 color in Aspose.Cells workbook | update Theme.ColorScheme for a chart
// Developer Intent: Apply custom RGB values to individual theme colors so that a chart inherits the new palette when the workbook is generated.
// Use Cases: Align chart colors with corporate branding by setting Accent1 and Accent2 to brand-specific shades. | Improve accessibility of generated reports by adjusting Text1 and Background1 contrast colors. | Produce localized versions of a spreadsheet with region‑specific theme palettes without manual editing.
// AI Prompts: Show C# code that changes Accent1, Accent2, Text1, and Background1 theme colors for a chart using Aspose.Cells. | Explain how to retrieve and edit the Theme.ColorScheme of a workbook after it has been created. | Provide a step‑by‑step guide to apply custom RGB values to theme colors so existing charts update automatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, and customize specific theme colors (Accent1, Accent2, Text1, Background1) using the Workbook.SetThemeColor method, which updates the chart's Theme.ColorScheme before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the series
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Modify theme colors that affect the chart
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 128, 128));      // teal
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 165, 0));    // orange
                workbook.SetThemeColor(ThemeColorType.Text1, Color.FromArgb(64, 64, 64));       // dark gray
                workbook.SetThemeColor(ThemeColorType.Background1, Color.FromArgb(255, 255, 200)); // light yellow

                // Save the workbook
                string outputPath = "ChartWithCustomThemeColors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

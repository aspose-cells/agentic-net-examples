// Title: Validate Automatic Chart Series Color Update After Changing Workbook Theme with Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds sample data and a column chart, records the initial Accent1 theme color and the first series color, changes Accent1 to LimeGreen using SetThemeColor, then checks that the series color reflects the new accent and saves the file for verification.
// Keywords: Aspose.Cells | .NET | C# | chart series color | theme accent | SetThemeColor | GetThemeColor | Accent1 | automatic color update | validate chart theme | Workbook theme change
// Common Searches: Aspose.Cells verify chart colors after theme change | C# set workbook theme accent and update chart series | how to test automatic chart color update in Aspose.Cells | GetThemeColor SetThemeColor example .NET | chart series follows theme accent Aspose.Cells
// Developer Intent: Confirm that existing chart series automatically adopt a new theme accent after calling SetThemeColor on the workbook.
// Use Cases: Create a report workbook, apply a custom theme, and ensure all chart series instantly match the new accent colors. | Automated unit test that iterates through every chart in a workbook to assert series colors update when the theme accent is modified. | Dynamic branding scenario where changing the workbook theme updates chart colors without rebuilding the charts.
// AI Prompts: Generate C# code that asserts each chart series color equals the current Accent1 theme color after calling SetThemeColor in Aspose.Cells. | Provide a loop that scans all charts in a workbook and verifies series colors update when the theme accent is changed. | Explain Aspose.Cells' mechanism for propagating theme color changes to existing chart series and whether a refresh call is required.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeUpdateValidation
{
    // This C# example creates a workbook, adds sample data and a column chart, records the initial Accent1 theme color and the first series color, changes Accent1 to LimeGreen using SetThemeColor, then checks that the series color reflects the new accent and saves the file for verification.
    public class ValidateChartSeriesThemeColors
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Populate sample data for the chart
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["A2"].PutValue("Jan");
                ws.Cells["A3"].PutValue("Feb");
                ws.Cells["A4"].PutValue("Mar");

                ws.Cells["B1"].PutValue("Series1");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["B3"].PutValue(20);
                ws.Cells["B4"].PutValue(30);

                ws.Cells["C1"].PutValue("Series2");
                ws.Cells["C2"].PutValue(15);
                ws.Cells["C3"].PutValue(25);
                ws.Cells["C4"].PutValue(35);

                // Add a column chart
                int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = ws.Charts[chartIdx];

                // Set the data range for the series
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Capture the initial theme accent color (Accent1) and the series color
                Color initialAccent = workbook.GetThemeColor(ThemeColorType.Accent1);
                Color initialSeriesColor = chart.NSeries[0].Area.ForegroundColor;

                Console.WriteLine($"Initial Accent1 Theme Color: {initialAccent}");
                Console.WriteLine($"Initial Series[0] Foreground Color: {initialSeriesColor}");

                // Change the theme Accent1 color to a distinct value (e.g., LimeGreen)
                Color newAccent = Color.LimeGreen;
                workbook.SetThemeColor(ThemeColorType.Accent1, newAccent);

                // After theme change, retrieve the series color again
                Color updatedSeriesColor = chart.NSeries[0].Area.ForegroundColor;
                Color updatedAccent = workbook.GetThemeColor(ThemeColorType.Accent1);

                Console.WriteLine($"Updated Accent1 Theme Color: {updatedAccent}");
                Console.WriteLine($"Updated Series[0] Foreground Color: {updatedSeriesColor}");

                // Validate that the series color reflects the new theme accent
                bool isUpdated = updatedSeriesColor.ToArgb() == newAccent.ToArgb();
                Console.WriteLine($"Series color updated to new theme accent: {isUpdated}");

                // Save the workbook (validation result can be inspected in the file)
                workbook.Save("ChartSeriesThemeValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateChartSeriesThemeColors.Run();
        }
    }
}

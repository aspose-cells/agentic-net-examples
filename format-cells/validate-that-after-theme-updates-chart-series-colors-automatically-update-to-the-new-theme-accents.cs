// Title: C# – Verify Chart Series Colors Auto‑Update After Workbook Theme Accent Change (Aspose.Cells)
// Description: Demonstrates how to create a workbook with a column chart, capture the initial Accent1 theme color and the first series foreground color, change Accent1 to Magenta using SetThemeColor, and confirm that the series color automatically reflects the new theme. The workbook is saved for visual verification.
// Keywords: Aspose.Cells | C# chart theme color | SetThemeColor | GetThemeColor | chart series color update | theme accent propagation | .NET Excel automation | validate chart colors after theme change
// Common Searches: Aspose.Cells change theme accent programmatically | chart series color follows workbook theme | C# verify theme color propagation to charts | GetThemeColor SetThemeColor example | auto‑update chart colors after theme change
// Developer Intent: Confirm that existing chart series automatically adopt a new workbook theme accent when the theme is modified via Aspose.Cells.
// Use Cases: Programmatically modify the workbook’s Accent1 color and ensure the first column‑chart series updates its foreground color without recreating the chart. | Log before‑and‑after theme and series colors to validate automatic theme propagation in automated tests. | Generate an Excel file where the visual appearance of charts reflects a custom theme applied at runtime.
// AI Prompts: Generate C# code using Aspose.Cells that changes the workbook theme accent and asserts that chart series colors update automatically. | Create a .NET unit test that builds a chart, records its series color, changes the theme accent, and verifies the series color matches the new accent. | Explain how GetThemeColor and SetThemeColor influence chart series colors in Aspose.Cells with step‑by‑step details.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeValidation
{
    // Demonstrates how to create a workbook with a column chart, capture the initial Accent1 theme color and the first series foreground color, change Accent1 to Magenta using SetThemeColor, and confirm that the series color automatically reflects the new theme. The workbook is saved for visual verification.
    public class ChartSeriesThemeUpdateDemo
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

                // Capture the initial theme accent color (Accent1) and the series' foreground color
                Color initialAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
                Color initialSeriesColor = chart.NSeries[0].Area.ForegroundColor;

                Console.WriteLine($"Initial Theme Accent1 Color : {initialAccent1}");
                Console.WriteLine($"Initial Series[0] Foreground Color : {initialSeriesColor}");

                // Update the theme: change Accent1 to a distinct color (Magenta)
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.Magenta);

                // After theme change, retrieve the updated theme accent color
                Color updatedAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
                // The series color should reflect the new theme accent automatically
                Color updatedSeriesColor = chart.NSeries[0].Area.ForegroundColor;

                Console.WriteLine($"Updated Theme Accent1 Color : {updatedAccent1}");
                Console.WriteLine($"Updated Series[0] Foreground Color : {updatedSeriesColor}");

                // Simple validation: check if the series color matches the updated accent color
                bool colorsMatch = updatedSeriesColor.ToArgb() == updatedAccent1.ToArgb();
                Console.WriteLine($"Series color reflects updated theme: {colorsMatch}");

                // Save the workbook to verify the visual result
                workbook.Save("ChartSeriesThemeUpdateDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

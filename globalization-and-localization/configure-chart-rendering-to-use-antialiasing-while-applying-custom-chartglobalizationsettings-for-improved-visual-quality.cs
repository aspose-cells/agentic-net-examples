// Title: Render Aspose.Cells Chart with Anti‑Aliasing and Custom ChartGlobalizationSettings (C#)
// Description: This example creates a workbook, adds quarterly sales data, builds a column chart, sets the value axis display unit to Hundreds, and applies a custom ChartGlobalizationSettings class that overrides axis unit names. The chart is rendered to a PNG file using ImageOrPrintOptions (which provides anti‑aliasing by default) and the workbook is saved.
// Keywords: Aspose.Cells | C# | chart anti-aliasing | ChartGlobalizationSettings | custom axis unit names | display unit label | render chart to PNG | ImageOrPrintOptions | localization | globalization
// Common Searches: Aspose.Cells enable anti aliasing for chart image | custom ChartGlobalizationSettings example C# | override axis unit name Aspose.Cells | set display unit label on chart Aspose.Cells | render high quality chart PNG Aspose.Cells
// Developer Intent: Generate a smooth‑edge chart image while applying a custom ChartGlobalizationSettings to localize axis unit labels.
// Use Cases: Create publication‑ready PNG charts with anti‑aliased rendering for reports and presentations. | Localize chart axis unit names for different cultures by subclassing ChartGlobalizationSettings. | Automate reporting pipelines that require both a high‑quality chart image and the original workbook. | Display custom unit labels on chart axes for financial or scientific dashboards.
// AI Prompts: Provide C# code to render an Aspose.Cells chart to PNG with anti‑aliasing and a custom ChartGlobalizationSettings. | Show how to subclass ChartGlobalizationSettings and override GetAxisUnitName for localized axis labels in Aspose.Cells. | Explain how to combine ImageOrPrintOptions smoothing with custom chart globalization to produce high‑quality, localized chart images.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartAntiAlias
{
    // Custom globalization settings for charts
    // This example creates a workbook, adds quarterly sales data, builds a column chart, sets the value axis display unit to Hundreds, and applies a custom ChartGlobalizationSettings class that overrides axis unit names. The chart is rendered to a PNG file using ImageOrPrintOptions (which provides anti‑aliasing by default) and the workbook is saved.
    public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Override axis unit names to demonstrate custom localization
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "Hundreds (custom)";
                case DisplayUnitType.Thousands:
                    return "Thousands (custom)";
                case DisplayUnitType.TenThousands:
                    return "Ten Thousands (custom)";
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);
                sheet.Cells["B4"].PutValue(180);

                // ---------- Add a column chart ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Quarterly Sales";

                // Set a display unit to see the custom axis unit name in action
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
                chart.ValueAxis.IsDisplayUnitLabelShown = true;

                // ---------- Apply custom chart globalization settings ----------
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new CustomChartGlobalizationSettings()
                };

                // ---------- Configure rendering options ----------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // No explicit anti‑aliasing properties are required; default rendering provides good quality.
                };

                // ---------- Render the chart to an image ----------
                string imagePath = "QuarterlySalesChart.png";
                chart.ToImage(imagePath, renderOptions);
                Console.WriteLine($"Chart rendered to image: {imagePath}");

                // ---------- Save the workbook ----------
                string workbookPath = "QuarterlySalesChart.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

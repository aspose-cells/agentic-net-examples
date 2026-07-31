// Title: C# – Override ChartGlobalizationSettings in Aspose.Cells to Localize Chart Legends to Arabic
// Description: Demonstrates how to create a custom ArabicChartGlobalizationSettings class that overrides legend‑related methods, assign it to Workbook.Settings.GlobalizationSettings.ChartSettings, and generate a column chart with Arabic category, value, and legend text using Aspose.Cells.
// Keywords: Aspose.Cells Arabic chart | ChartGlobalizationSettings override | C# chart legend localization | Arabic legend text Aspose.Cells | GetLegendIncreaseName Arabic | custom chart globalization .NET | Excel Arabic legends | Aspose.Cells chart localization example | globalization settings chart Arabic
// Common Searches: how to translate chart legends to Arabic in Aspose.Cells | override ChartGlobalizationSettings for Arabic legends | Aspose.Cells Arabic chart example C# | set Arabic legend names in Excel charts programmatically | customize chart globalization settings Aspose.Cells
// Developer Intent: Provide Arabic translations for all legend entries in charts generated with Aspose.Cells by overriding ChartGlobalizationSettings.
// Use Cases: Generate sales or financial charts where legend items like Increase, Decrease, and Total appear in Arabic. | Apply a single ArabicChartGlobalizationSettings instance to a workbook so every new chart automatically uses Arabic terminology. | Create multilingual Excel reports that require Arabic chart legends without manually editing each chart.
// AI Prompts: Write C# code that overrides ChartGlobalizationSettings to supply French translations for chart legends in Aspose.Cells. | Show how to attach a custom ChartGlobalizationSettings object to an existing workbook that already contains multiple charts. | Explain how to extend the ArabicChartGlobalizationSettings class to customize series names and other chart labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsArabicChartDemo
{
    // Custom globalization settings that provide Arabic translations for legend entries
    // Demonstrates how to create a custom ArabicChartGlobalizationSettings class that overrides legend‑related methods, assign it to Workbook.Settings.GlobalizationSettings.ChartSettings, and generate a column chart with Arabic category, value, and legend text using Aspose.Cells.
    public class ArabicChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Arabic text for "Increase" in legend
        public override string GetLegendIncreaseName()
        {
            return "زيادة";
        }

        // Arabic text for "Decrease" in legend
        public override string GetLegendDecreaseName()
        {
            return "نقصان";
        }

        // Arabic text for "Total" in legend
        public override string GetLegendTotalName()
        {
            return "المجموع";
        }

        // Optionally override other methods if needed (e.g., series name, other label)
        public override string GetSeriesName()
        {
            return "سلسلة";
        }

        public override string GetOtherName()
        {
            return "أخرى";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply the custom Arabic chart globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ArabicChartGlobalizationSettings()
            };

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("الفئة");   // "Category" in Arabic
            sheet.Cells["A2"].PutValue("المنتج 1");
            sheet.Cells["A3"].PutValue("المنتج 2");
            sheet.Cells["B1"].PutValue("القيمة"); // "Value" in Arabic
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(250);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Bind data to the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set chart title (optional)
            chart.Title.Text = "مخطط المبيعات";

            // Save the workbook
            workbook.Save("ArabicChartDemo.xlsx");
        }
    }
}

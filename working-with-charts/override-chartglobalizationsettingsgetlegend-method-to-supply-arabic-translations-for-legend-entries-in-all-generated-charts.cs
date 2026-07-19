// Title: Aspose.Cells C# – Override ChartGlobalizationSettings to render Arabic legend and axis text
// Description: This example shows how to create a custom class derived from SettableChartGlobalizationSettings that returns Arabic strings for legend entries, chart title, axis titles, and generic series names. The custom settings are assigned to workbook.Settings.GlobalizationSettings.ChartSettings, a column chart with Arabic data is built, and the workbook is saved so that all chart UI elements appear in Arabic.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | Arabic chart legend | .NET | C# | SettableChartGlobalizationSettings | chart localization | right-to-left chart | globalization settings example | GitHub Aspose.Cells Arabic
// Common Searches: How to localize Aspose.Cells chart legends to Arabic | Override ChartGlobalizationSettings for custom legend names in .NET | Arabic axis titles Aspose.Cells C# | Set Arabic chart title and series name with Aspose.Cells | Aspose.Cells Arabic globalization settings sample
// Developer Intent: Provide Arabic text for every chart UI element by customizing ChartGlobalizationSettings.
// Use Cases: Generate financial or sales reports for Middle‑East audiences with fully Arabic chart legends and titles. | Apply the same ArabicChartGlobalizationSettings to multiple workbooks to maintain consistent localization. | Combine Arabic chart globalization with right‑to‑left worksheet layout for complete Arabic UI support.
// AI Prompts: Write C# code that overrides SettableChartGlobalizationSettings to return French translations for chart legends in Aspose.Cells. | Demonstrate how to attach a custom ChartGlobalizationSettings object to an existing workbook that already contains charts. | Explain how to integrate ArabicChartGlobalizationSettings with worksheet.RightToLeft = true for proper Arabic rendering.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsArabicChartDemo
{
    // Custom globalization settings that provide Arabic translations for legend entries
    // This example shows how to create a custom class derived from SettableChartGlobalizationSettings that returns Arabic strings for legend entries, chart title, axis titles, and generic series names. The custom settings are assigned to workbook.Settings.GlobalizationSettings.ChartSettings, a column chart with Arabic data is built, and the workbook is saved so that all chart UI elements appear in Arabic.
    public class ArabicChartGlobalizationSettings : SettableChartGlobalizationSettings
    {
        // Arabic text for the "Increase" legend entry
        public override string GetLegendIncreaseName()
        {
            return "زيادة";
        }

        // Arabic text for the "Decrease" legend entry
        public override string GetLegendDecreaseName()
        {
            return "نقصان";
        }

        // Arabic text for the "Total" legend entry
        public override string GetLegendTotalName()
        {
            return "الإجمالي";
        }

        // Arabic text for generic series name
        public override string GetSeriesName()
        {
            return "سلسلة";
        }

        // Arabic text for chart title
        public override string GetChartTitleName()
        {
            return "عنوان المخطط";
        }

        // Arabic text for other chart elements
        public override string GetOtherName()
        {
            return "أخرى";
        }

        // Arabic text for axis title
        public override string GetAxisTitleName()
        {
            return "عنوان المحور";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart (Arabic labels)
                sheet.Cells["A1"].PutValue("الفئة");
                sheet.Cells["A2"].PutValue("الفئة 1");
                sheet.Cells["A3"].PutValue("الفئة 2");
                sheet.Cells["B1"].PutValue("القيمة");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["B3"].PutValue(250);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";
                chart.Title.Text = "مخطط تجريبي";

                // Apply the custom Arabic globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ArabicChartGlobalizationSettings()
                };

                // Save the workbook; the chart will display Arabic legend entries
                string outputPath = "ArabicChartDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

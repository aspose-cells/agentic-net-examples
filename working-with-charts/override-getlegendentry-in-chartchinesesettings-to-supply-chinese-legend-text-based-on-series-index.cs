// Title: Custom Chinese Legend Text for Aspose.Cells Charts via GetLegendEntry Override (C#)
// Description: Shows how to subclass SettableChartGlobalizationSettings in Aspose.Cells for .NET, implement a GetLegendEntry method that returns predefined Chinese series names (or a generic fallback) and assign those names to chart series, resulting in a column chart with localized legend entries saved as ChartChineseSettings.xlsx.
// Keywords: Aspose.Cells | C# | chart legend localization | Chinese legend names | SettableChartGlobalizationSettings | GetLegendEntry | custom globalization | column chart | series name assignment | .NET workbook | multilingual charts
// Common Searches: Aspose.Cells set chart legend language | C# custom chart legend text Aspose | how to localize chart legends in Aspose.Cells | GetLegendEntry override example | SettableChartGlobalizationSettings Chinese names | Aspose.Cells assign series name programmatically | create chart with Chinese legend Aspose.Cells
// Developer Intent: Create a subclass of SettableChartGlobalizationSettings that provides Chinese legend entries based on series index and apply it to a chart.
// Use Cases: Generate column charts with Chinese legend entries for each data series. | Localize chart legends in multilingual workbooks by mapping series indices to language‑specific names. | Provide a fallback series name when the index exceeds predefined Chinese names to ensure all legends appear. | Integrate custom globalization settings into existing Aspose.Cells chart generation pipelines. | Automate translation of chart legends without modifying source data.
// AI Prompts: Write a C# class that extends SettableChartGlobalizationSettings and implements GetLegendEntry to return Chinese names for known series indices and a generic name otherwise. | Demonstrate how to assign custom legend text to each series of an Aspose.Cells chart using a ChartChineseSettings instance. | Explain how to expand the GetLegendEntry method to support additional languages or dynamic lookup tables for chart legend localization.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provides Chinese legend names.
    // Shows how to subclass SettableChartGlobalizationSettings in Aspose.Cells for .NET, implement a GetLegendEntry method that returns predefined Chinese series names (or a generic fallback) and assign those names to chart series, resulting in a column chart with localized legend entries saved as ChartChineseSettings.xlsx.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        // Override (or implement) a method that returns the legend text for a given series index.
        // The base class does not define GetLegendEntry, so we provide our own implementation.
        public string GetLegendEntry(int index)
        {
            // Define Chinese names for the first few series.
            string[] chineseNames = { "第一系列", "第二系列", "第三系列", "第四系列", "第五系列" };

            // Return the matching name if within range; otherwise generate a generic name.
            if (index >= 0 && index < chineseNames.Length)
                return chineseNames[index];

            return $"系列{index + 1}";
        }
    }

    public class ChartChineseSettingsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(40);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories.
            chart.NSeries.Add("B2:C3", true);          // Two series.
            chart.NSeries.CategoryData = "A2:A3";      // Categories.

            // Apply Chinese legend names using the custom settings class.
            ChartChineseSettings chineseSettings = new ChartChineseSettings();

            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                // The series name is reflected in the legend entry.
                chart.NSeries[i].Name = chineseSettings.GetLegendEntry(i);
            }

            // Save the workbook.
            workbook.Save("ChartChineseSettings.xlsx");
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            ChartChineseSettingsDemo.Run();
        }
    }
}

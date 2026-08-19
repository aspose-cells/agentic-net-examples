// Title: Override GetLegendEntry to Supply Chinese Legend Text for Chart Series in Aspose.Cells (.NET)
// Description: Demonstrates a ChartChineseSettings class that overrides GetLegendEntry to return Chinese series names from a predefined array or a fallback pattern, creates a workbook with sample data, builds a column chart, retrieves localized legend entries for each series, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | Chart legend localization | Chinese legend text | GetLegendEntry override | custom chart legend | Excel chart series | column chart Aspose | i18n Aspose.Cells | chart series naming
// Common Searches: Aspose.Cells customize chart legend text | GetLegendEntry Chinese series names Aspose | set Chinese legend for Excel chart .NET | override chart legend method Aspose.Cells | localize chart legends in C# Excel library
// Developer Intent: Implement GetLegendEntry so that each chart series receives a Chinese legend label based on its index, with a fallback for undefined indexes.
// Use Cases: Create Excel reports with column charts that display legend entries in Chinese. | Provide automatic fallback Chinese names when the series count exceeds a predefined list. | Integrate ChartChineseSettings into automated reporting pipelines to ensure localized chart legends before workbook export.
// AI Prompts: Write a C# GetLegendEntry method that returns Chinese legend strings for chart series using an array and a fallback format. | Show how to apply the custom GetLegendEntry method to assign localized legend text to each series in an Aspose.Cells chart before saving. | Explain how to extend ChartChineseSettings to support additional languages while keeping index‑based legend retrieval.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates a ChartChineseSettings class that overrides GetLegendEntry to return Chinese series names from a predefined array or a fallback pattern, creates a workbook with sample data, builds a column chart, retrieves localized legend entries for each series, and saves the file as an Excel workbook.
public class ChartChineseSettings
{
    /// <summary>
    /// Returns the Chinese legend text for a given series index.
    /// </summary>
    public string GetLegendEntry(int seriesIndex)
    {
        string[] chineseNames = { "系列一", "系列二", "系列三", "系列四", "系列五" };

        if (seriesIndex >= 0 && seriesIndex < chineseNames.Length)
            return chineseNames[seriesIndex];

        // Fallback for indexes beyond the predefined array.
        return $"系列{seriesIndex + 1}";
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("第一季度");
            sheet.Cells["A3"].PutValue("第二季度");
            sheet.Cells["B1"].PutValue("Series A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["C1"].PutValue("Series B");
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(130);

            // Add a column chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series to the chart.
            chart.NSeries.Add("B2:B3", true); // Series A
            chart.NSeries.Add("C2:C3", true); // Series B
            chart.NSeries.CategoryData = "A2:A3";

            // Instantiate the custom Chinese settings.
            ChartChineseSettings chineseSettings = new ChartChineseSettings();

            // Retrieve and display Chinese legend names for each series.
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                string chineseLegend = chineseSettings.GetLegendEntry(i);
                Console.WriteLine($"Series {i} Chinese Legend: {chineseLegend}");
            }

            // Save the workbook.
            string outputPath = "ChartWithChineseLegends.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

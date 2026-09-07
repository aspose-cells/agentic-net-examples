// Title: How to Localize Waterfall Chart Legend Entries (Increase, Decrease, Total) to French with Aspose.Cells in C#
// AI Prompts: Write C# code that creates a Waterfall chart in an Aspose.Cells workbook and uses SettableChartGlobalizationSettings to rename the legend items "Increase", "Decrease", and "Total" to French terms. | Show how to retrieve and print the French legend names with GetLegendIncreaseName, GetLegendDecreaseName, and GetLegendTotalName before saving the workbook.
// Common Searches: Aspose.Cells set French legend text for Waterfall chart C# example | C# chart globalization settings change legend labels to French | How to customize Waterfall chart legend language using Aspose.Cells | Verify localized legend names in Aspose.Cells workbook before saving
// Tags: chart legend localization with Aspose.Cells | Aspose.Cells French legend translation | C# globalization settings for Excel charts | customize legend entries in .NET workbook | export workbook with localized chart labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a Waterfall chart, and applies SettableChartGlobalizationSettings to replace the default legend entries (Increase, Decrease, Total) with French equivalents. It prints the French names to the console for verification and saves the file as 'FrenchLegendChart.xlsx'.
class LocalizeLegendFrench
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Mois");
        worksheet.Cells["A2"].PutValue("Janvier");
        worksheet.Cells["A3"].PutValue("Février");
        worksheet.Cells["A4"].PutValue("Mars");
        worksheet.Cells["B1"].PutValue("Valeur");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        // Create a Waterfall chart (legend uses Increase/Decrease/Total)
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 6, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Ventes Mensuelles";

        // Configure French legend labels using SettableChartGlobalizationSettings
        SettableChartGlobalizationSettings frenchSettings = new SettableChartGlobalizationSettings();
        frenchSettings.SetLegendIncreaseName("Augmentation");
        frenchSettings.SetLegendDecreaseName("Diminution");
        frenchSettings.SetLegendTotalName("Total");

        // Apply the custom globalization settings to the workbook (lifecycle rule)
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = frenchSettings
        };

        // Verify that the settings return the French strings
        Console.WriteLine("Legend Increase (French): " + frenchSettings.GetLegendIncreaseName());
        Console.WriteLine("Legend Decrease (French): " + frenchSettings.GetLegendDecreaseName());
        Console.WriteLine("Legend Total (French): " + frenchSettings.GetLegendTotalName());

        // Save the workbook (save rule)
        workbook.Save("FrenchLegendChart.xlsx");
    }
}

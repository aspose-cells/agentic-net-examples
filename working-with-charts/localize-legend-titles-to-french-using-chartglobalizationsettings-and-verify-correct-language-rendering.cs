// Title: Localize Excel chart legends to French with Aspose.Cells C# (ChartGlobalizationSettings)
// Description: Creates a workbook, adds a column chart with monthly sales data, configures SettableChartGlobalizationSettings with French legend and title strings, applies the settings to the workbook, prints the localized values for verification, and saves the file as LocalizedChartLegendFrench.xlsx.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | SettableChartGlobalizationSettings | French localization | Excel chart legend translation | chart title French | globalization settings example | Aspose.Cells chart localization | Excel French UI
// Common Searches: Aspose.Cells change chart legend language to French | C# set chart legend French Aspose.Cells | ChartGlobalizationSettings French example | How to localize Excel chart titles with Aspose | SettableChartGlobalizationSettings usage in .NET
// Developer Intent: The developer wants to replace default English legend labels and chart titles with French equivalents using Aspose.Cells’ ChartGlobalizationSettings and confirm that the changes are applied correctly.
// Use Cases: Define French terms for legend increase, decrease, total, and apply them globally to all charts in a workbook. | Set a French chart title and series name once, affecting every chart without per‑chart code changes. | Output the localized strings to the console to validate the settings before saving the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to localize chart legends and titles to German with SettableChartGlobalizationSettings. | Explain how to read, modify, and reapply existing ChartGlobalizationSettings after loading an Excel file in Aspose.Cells. | Provide a step‑by‑step method to programmatically verify that French legend strings appear correctly in the saved Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart with monthly sales data, configures SettableChartGlobalizationSettings with French legend and title strings, applies the settings to the workbook, prints the localized values for verification, and saves the file as LocalizedChartLegendFrench.xlsx.
class LocalizeChartLegendFrench
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Mois");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Fév");
        worksheet.Cells["A4"].PutValue("Mar");
        worksheet.Cells["B1"].PutValue("Ventes");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories
        chart.Title.Text = "Ventes Mensuelles";

        // Create a SettableChartGlobalizationSettings instance with French texts
        SettableChartGlobalizationSettings frenchSettings = new SettableChartGlobalizationSettings();
        frenchSettings.SetLegendIncreaseName("Augmentation");
        frenchSettings.SetLegendDecreaseName("Diminution");
        frenchSettings.SetLegendTotalName("Total");
        frenchSettings.SetChartTitleName("Titre du graphique");
        frenchSettings.SetSeriesName("Série");
        frenchSettings.SetOtherName("Autre");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = frenchSettings
        };

        // Verify that the settings return the French strings
        Console.WriteLine("Legend Increase (French): " + frenchSettings.GetLegendIncreaseName());
        Console.WriteLine("Legend Decrease (French): " + frenchSettings.GetLegendDecreaseName());
        Console.WriteLine("Legend Total (French): " + frenchSettings.GetLegendTotalName());
        Console.WriteLine("Chart Title (French): " + frenchSettings.GetChartTitleName());

        // Save the workbook
        workbook.Save("LocalizedChartLegendFrench.xlsx");
    }
}

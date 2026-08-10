// Title: Aspose.Cells .NET – Localize Excel Chart Legend Titles to French with SettableChartGlobalizationSettings
// Description: Creates a workbook with monthly sales data, adds a column chart, and uses SettableChartGlobalizationSettings to replace the default legend entries (Increase, Decrease, Total) with French terms. The code prints the localized strings for verification and saves the file as LocalizedChartLegendFrench.xlsx.
// Keywords: Aspose.Cells | chart legend localization | French legend labels | SettableChartGlobalizationSettings | .NET Excel chart | globalization settings | Excel chart translation | column chart French | Aspose.Cells API | Excel automation French
// Common Searches: Aspose.Cells change chart legend language to French | SettableChartGlobalizationSettings example .NET | localize Excel chart legend labels | French legend names for Aspose.Cells chart | globalize chart legends in C#
// Developer Intent: Apply French text to chart legend entries (Increase, Decrease, Total) using Aspose.Cells globalization settings and confirm the changes are persisted in the saved workbook.
// Use Cases: Generate sales reports for French‑speaking audiences with correctly translated chart legends. | Standardize legend terminology across multiple charts in a single workbook by reusing a French globalization object. | Automate verification of localized legend strings before distributing the Excel file.
// AI Prompts: Write C# code that localizes chart legend titles to German using Aspose.Cells SettableChartGlobalizationSettings and prints the results. | Explain how to apply a custom SettableChartGlobalizationSettings instance to all charts in an existing workbook with Aspose.Cells for .NET. | Create a unit test in MSTest that opens the saved Excel file and asserts that the legend entries are "Augmentation", "Diminution", and "Total".

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with monthly sales data, adds a column chart, and uses SettableChartGlobalizationSettings to replace the default legend entries (Increase, Decrease, Total) with French terms. The code prints the localized strings for verification and saves the file as LocalizedChartLegendFrench.xlsx.
class LocalizeChartLegendFrench
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Mois");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Fév");
        worksheet.Cells["A4"].PutValue("Mar");
        worksheet.Cells["B1"].PutValue("Ventes");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        // Create a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Ventes Mensuelles";

        // Create customizable globalization settings for French legend labels
        SettableChartGlobalizationSettings frenchSettings = new SettableChartGlobalizationSettings();
        frenchSettings.SetLegendIncreaseName("Augmentation");
        frenchSettings.SetLegendDecreaseName("Diminution");
        frenchSettings.SetLegendTotalName("Total");

        // Apply the settings to the workbook's globalization settings
        workbook.Settings.GlobalizationSettings.ChartSettings = frenchSettings;

        // Verify that the settings contain the French strings
        Console.WriteLine("Legend Increase (French): " + frenchSettings.GetLegendIncreaseName());
        Console.WriteLine("Legend Decrease (French): " + frenchSettings.GetLegendDecreaseName());
        Console.WriteLine("Legend Total (French): " + frenchSettings.GetLegendTotalName());

        // Save the workbook
        workbook.Save("LocalizedChartLegendFrench.xlsx");
    }
}

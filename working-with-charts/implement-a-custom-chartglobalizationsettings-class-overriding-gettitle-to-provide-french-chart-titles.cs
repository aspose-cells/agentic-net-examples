// Title: Custom ChartGlobalizationSettings in Aspose.Cells (.NET) – French Chart Title
// Description: Demonstrates how to subclass ChartGlobalizationSettings, override GetChartTitleName to return the French string "Titre du graphique", assign the custom settings to a workbook, create sample data, add a column chart, apply the localized title, and save the file as CustomChartGlobalization.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | custom chart localization | French chart title | .NET | C# Excel automation | Excel chart globalization | override GetChartTitleName | multilingual reporting | Excel chart title translation
// Common Searches: Aspose.Cells chart globalization example | how to localize chart titles in Aspose.Cells C# | override ChartGlobalizationSettings for French | custom chart title with Aspose.Cells .NET | Excel chart title translation using Aspose
// Developer Intent: Create a reusable ChartGlobalizationSettings subclass that supplies French titles for Excel charts generated with Aspose.Cells.
// Use Cases: Provide French‑language chart titles in automated Excel reports without hard‑coding strings. | Swap different localization classes (e.g., German, Spanish) to produce multilingual workbooks from the same chart‑creation logic. | Standardize chart labeling across an enterprise reporting suite by centralizing globalization settings.
// AI Prompts: Write C# code that defines a CustomChartGlobalizationSettings class returning a German chart title and applies it to an Aspose.Cells workbook. | Explain how to make GetChartTitleName return dynamic titles based on a workbook's locale property. | Give step‑by‑step instructions to verify that the French chart title appears correctly in the saved Excel file.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Custom globalization settings that provide French chart titles
// Demonstrates how to subclass ChartGlobalizationSettings, override GetChartTitleName to return the French string "Titre du graphique", assign the custom settings to a workbook, create sample data, add a column chart, apply the localized title, and save the file as CustomChartGlobalization.xlsx using Aspose.Cells for .NET.
public class CustomChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Override the method that returns the default chart title name
    public override string GetChartTitleName()
    {
        // French translation for "Chart Title"
        return "Titre du graphique";
    }
}

public class ChartGlobalizationDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new CustomChartGlobalizationSettings()
        };

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Catégorie");
        sheet.Cells["A2"].PutValue("Janvier");
        sheet.Cells["A3"].PutValue("Février");
        sheet.Cells["B1"].PutValue("Valeur");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Set the chart title using the French title from the custom globalization settings
        chart.Title.Text = workbook.Settings.GlobalizationSettings.ChartSettings.GetChartTitleName();

        // Save the workbook to a file
        workbook.Save("CustomChartGlobalization.xlsx");
    }
}

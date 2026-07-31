// Title: Custom French Chart Title Using ChartGlobalizationSettings in Aspose.Cells for .NET
// Description: Demonstrates how to subclass ChartGlobalizationSettings, override GetChartTitleName to return a French title, assign the custom settings to a workbook, create a column chart with sample data, and save the file as FrenchChartTitle.xlsx.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | French chart title | C# | .NET | Excel chart localization | GetChartTitleName override | custom globalization settings | multilingual Excel reports | chart title translation
// Common Searches: Aspose.Cells how to localize chart titles | override ChartGlobalizationSettings for French title | C# example of chart globalization in Aspose.Cells | set workbook globalization settings Aspose.Cells | custom chart title language Aspose.Cells .NET
// Developer Intent: Create a reusable ChartGlobalizationSettings subclass that automatically provides French titles for all charts in an Aspose.Cells workbook.
// Use Cases: Apply FrenchChartGlobalizationSettings to a workbook so every chart displays the French title without manual Title.Text changes. | Build multilingual Excel reports by defining separate ChartGlobalizationSettings subclasses for each target language and swapping them at runtime. | Standardize chart UI across an organization by centralizing title localization in a custom globalization class.
// AI Prompts: Generate C# code that implements a GermanChartGlobalizationSettings class to provide German chart titles in Aspose.Cells. | List all overridable methods of ChartGlobalizationSettings and explain how each can be used for chart UI localization. | Show how to switch between EnglishChartGlobalizationSettings and FrenchChartGlobalizationSettings for different worksheets in the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Custom globalization settings that provide a French chart title
// Demonstrates how to subclass ChartGlobalizationSettings, override GetChartTitleName to return a French title, assign the custom settings to a workbook, create a column chart with sample data, and save the file as FrenchChartTitle.xlsx.
public class FrenchChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Override the method that returns the chart title name
    public override string GetChartTitleName()
    {
        // French translation for "Chart Title"
        return "Titre du graphique";
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        worksheet.Cells["A1"].PutValue("Catégorie");
        worksheet.Cells["A2"].PutValue("Janvier");
        worksheet.Cells["A3"].PutValue("Février");
        worksheet.Cells["B1"].PutValue("Valeur");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Apply the custom French globalization settings for chart titles
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new FrenchChartGlobalizationSettings()
        };

        // Optionally set a placeholder title; when rendered it will use the French name
        chart.Title.Text = "Placeholder";

        // Save the workbook
        workbook.Save("FrenchChartTitle.xlsx");
    }
}

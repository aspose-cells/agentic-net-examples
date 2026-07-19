// Title: C# – Custom ChartGlobalizationSettings to Localize Aspose.Cells Chart Titles in French
// Description: Demonstrates how to create a FrenchChartGlobalizationSettings class that inherits from ChartGlobalizationSettings, overrides GetChartTitleName to return "Titre du graphique", assigns it to Workbook.Settings.GlobalizationSettings.ChartSettings, and applies the title to a column chart before saving the workbook as FrenchChartTitle.xlsx.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | C# | .NET | French localization | chart title translation | override GetChartTitleName | custom globalization | multilingual Excel reports | Excel chart title French
// Common Searches: Aspose.Cells override chart title language | C# custom ChartGlobalizationSettings example | French chart title Aspose.Cells | How to localize chart titles in Aspose.Cells .NET | Set workbook globalization settings for charts
// Developer Intent: Create a subclass of ChartGlobalizationSettings that returns a French chart title and apply it to a workbook.
// Use Cases: Produce Excel reports with French chart titles for European audiences | Switch chart language by swapping ChartGlobalizationSettings implementations at runtime | Maintain a library of localization classes for consistent multilingual chart rendering across projects
// AI Prompts: Generate a C# ChartGlobalizationSettings subclass that returns a Spanish chart title. | Show code to toggle between English and French chart globalization settings in an Aspose.Cells workbook. | Provide an example that localizes axis labels and legend text using a custom ChartGlobalizationSettings implementation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomChartGlobalization
{
    // Custom globalization settings for charts that returns French titles
    // Demonstrates how to create a FrenchChartGlobalizationSettings class that inherits from ChartGlobalizationSettings, overrides GetChartTitleName to return "Titre du graphique", assigns it to Workbook.Settings.GlobalizationSettings.ChartSettings, and applies the title to a column chart before saving the workbook as FrenchChartTitle.xlsx.
    public class FrenchChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Override the method that provides the chart title name
        public override string GetChartTitleName()
        {
            // French translation for "Chart Title"
            return "Titre du graphique";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new FrenchChartGlobalizationSettings()
            };

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Catégorie");
            sheet.Cells["A2"].PutValue("Janvier");
            sheet.Cells["A3"].PutValue("Février");
            sheet.Cells["A4"].PutValue("Mars");
            sheet.Cells["B1"].PutValue("Valeur");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use the overridden GetChartTitleName to set the chart title
            string frenchTitle = workbook.Settings.GlobalizationSettings.ChartSettings.GetChartTitleName();
            chart.Title.Text = frenchTitle;

            // Save the workbook
            workbook.Save("FrenchChartTitle.xlsx");
        }
    }
}

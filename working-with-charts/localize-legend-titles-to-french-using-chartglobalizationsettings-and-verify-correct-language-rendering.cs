// Title: C# – Localize Aspose.Cells Chart Legend to French with SettableChartGlobalizationSettings
// Description: Shows how to create a workbook, add a column chart with monthly sales data, and replace default legend strings with French terms using SettableChartGlobalizationSettings. The settings are applied, printed for verification, and the workbook is saved as LocalizedChart.xlsx to confirm visual rendering of French legends.
// Keywords: Aspose.Cells | C# | .NET | chart localization | French legend | SettableChartGlobalizationSettings | ChartGlobalizationSettings | Excel chart globalization | Aspose.Cells example | localize chart titles
// Common Searches: Aspose.Cells change chart legend language | Set chart legend to French Aspose.Cells | C# chart globalization French Aspose | How to localize chart titles in Aspose.Cells .NET | Apply custom chart globalization settings Aspose.Cells
// Developer Intent: Apply French language labels to chart legend and related UI elements with Aspose.Cells, then verify they appear correctly in the generated Excel file.
// Use Cases: Generate a sales column chart and localize legend, title, axis, and series names to French before saving. | Retrieve each localized string via Get methods for logging or debugging purposes. | Reuse the same SettableChartGlobalizationSettings instance across multiple workbooks or cultures. | Open the saved LocalizedChart.xlsx to visually confirm French legend rendering.
// AI Prompts: Write C# code that localizes all chart UI text to German using Aspose.Cells SettableChartGlobalizationSettings. | Create a function that accepts a culture code and returns a configured SettableChartGlobalizationSettings object for a workbook. | Explain how to programmatically verify that French legend labels are rendered correctly in the Excel file produced by Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLocalization
{
    // Shows how to create a workbook, add a column chart with monthly sales data, and replace default legend strings with French terms using SettableChartGlobalizationSettings. The settings are applied, printed for verification, and the workbook is saved as LocalizedChart.xlsx to confirm visual rendering of French legends.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Mois");
            sheet.Cells["A2"].PutValue("Janvier");
            sheet.Cells["A3"].PutValue("Février");
            sheet.Cells["A4"].PutValue("Mars");
            sheet.Cells["B1"].PutValue("Ventes");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Ventes Mensuelles";

            // Create a SettableChartGlobalizationSettings instance and set French labels
            SettableChartGlobalizationSettings frenchSettings = new SettableChartGlobalizationSettings();
            frenchSettings.SetLegendIncreaseName("Augmentation");
            frenchSettings.SetLegendDecreaseName("Diminution");
            frenchSettings.SetLegendTotalName("Total");
            frenchSettings.SetChartTitleName("Titre du graphique");
            frenchSettings.SetSeriesName("Série");
            frenchSettings.SetAxisTitleName("Titre de l'axe");
            frenchSettings.SetOtherName("Autre");

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings.ChartSettings = frenchSettings;

            // Verify that the settings return the expected French strings
            Console.WriteLine("Legend Increase (French): " + frenchSettings.GetLegendIncreaseName());
            Console.WriteLine("Legend Decrease (French): " + frenchSettings.GetLegendDecreaseName());
            Console.WriteLine("Legend Total (French): " + frenchSettings.GetLegendTotalName());
            Console.WriteLine("Chart Title (French): " + frenchSettings.GetChartTitleName());
            Console.WriteLine("Series Name (French): " + frenchSettings.GetSeriesName());
            Console.WriteLine("Axis Title (French): " + frenchSettings.GetAxisTitleName());
            Console.WriteLine("Other (French): " + frenchSettings.GetOtherName());

            // Save the workbook to verify visual rendering of the French legend labels
            workbook.Save("LocalizedChart.xlsx");
        }
    }
}

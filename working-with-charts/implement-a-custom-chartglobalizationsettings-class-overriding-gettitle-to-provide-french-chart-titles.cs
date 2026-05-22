using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Custom globalization settings that provide French chart titles
    public class FrenchChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Override the method that returns the default chart title name
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
            sheet.Cells["A2"].PutValue("Produit A");
            sheet.Cells["A3"].PutValue("Produit B");
            sheet.Cells["B1"].PutValue("Valeur");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(250);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set a custom title; the default title text (if not set) would use GetChartTitleName()
            chart.Title.Text = "Ventes Mensuelles";

            // Save the workbook
            workbook.Save("FrenchChartGlobalization.xlsx");
        }
    }
}
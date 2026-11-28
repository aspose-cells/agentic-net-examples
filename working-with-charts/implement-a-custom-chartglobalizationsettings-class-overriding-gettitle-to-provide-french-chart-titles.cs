// Title: Create a custom FrenchChartGlobalizationSettings class to localize chart titles and axis labels in Aspose.Cells for .NET
// Common Searches: how to set French titles for Aspose.Cells charts using ChartGlobalizationSettings in C# | example of overriding GetChartTitleName to localize Aspose.Cells chart text | custom ChartGlobalizationSettings class for French Excel chart localization | Aspose.Cells C# code to apply French axis labels to a column chart | changing series name to French in Aspose.Cells generated workbook
// Tags: custom ChartGlobalizationSettings implementation Aspose.Cells | French chart title localization Aspose.Cells .NET | override GetAxisTitleName for Excel chart axis localization | Aspose.Cells column chart with French localization | localize series name in Aspose.Cells chart using C# | Aspose.Cells chart globalization settings example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartGlobalizationDemo
{
    // Custom globalization settings that provide French titles for charts and axes.
    // Demonstrates how to subclass ChartGlobalizationSettings in C# to provide French strings for chart titles, axis titles, series names, and other labels, then applies these translations to a column chart in an Aspose.Cells workbook and saves it as an Excel file.
    public class FrenchChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // French translation for "Chart Title"
        public override string GetChartTitleName() => "Titre du graphique";

        // French translation for "Axis Title"
        public override string GetAxisTitleName() => "Titre de l'axe";

        // Additional French strings
        public override string GetOtherName() => "Autre";
        public override string GetSeriesName() => "Série";
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the chart.
                cells["A1"].PutValue("Catégorie");
                cells["A2"].PutValue("Produit A");
                cells["A3"].PutValue("Produit B");
                cells["A4"].PutValue("Produit C");

                cells["B1"].PutValue("Valeur");
                cells["B2"].PutValue(120);
                cells["B3"].PutValue(250);
                cells["B4"].PutValue(180);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply French titles using the custom globalization settings.
                FrenchChartGlobalizationSettings frSettings = new FrenchChartGlobalizationSettings();

                chart.Title.Text = frSettings.GetChartTitleName();
                chart.Title.IsVisible = true;

                chart.CategoryAxis.Title.Text = frSettings.GetAxisTitleName();
                chart.CategoryAxis.Title.IsVisible = true;

                chart.ValueAxis.Title.Text = frSettings.GetAxisTitleName();
                chart.ValueAxis.Title.IsVisible = true;

                // Configure legend (visibility is true by default).
                chart.Legend.Position = LegendPositionType.Right;

                // Customize series name which appears in the legend.
                chart.NSeries[0].Name = frSettings.GetSeriesName();

                // Save the workbook.
                string outputPath = "FrenchChartDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

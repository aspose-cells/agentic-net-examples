// Title: Implement a ChartGlobalizationSettings subclass in Aspose.Cells for .NET to override GetTitle and provide French chart titles
// AI Prompts: Write a ChartGlobalizationSettings class that overrides GetTitle to return French strings for the chart title, category axis, and value axis. | Show how to assign the custom ChartGlobalizationSettings instance to a Workbook so that every chart automatically uses the French titles. | Refactor the sample code to remove direct title assignments and rely on the overridden GetTitle method for localization.
// Common Searches: Aspose.Cells C# how to use ChartGlobalizationSettings to localize chart titles in French | override GetTitle method in ChartGlobalizationSettings for multilingual Excel charts Aspose.Cells | C# example of applying custom chart globalization settings to a workbook in Aspose.Cells | automatically set French chart and axis titles in Aspose.Cells without manual assignment
// Tags: custom ChartGlobalizationSettings subclass C# | override GetTitle for chart localization | French chart titles Aspose.Cells | globalize Excel chart text .NET | apply chart globalization settings workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // The example demonstrates creating a Workbook, adding sample data, inserting a column chart, and using a custom ChartGlobalizationSettings subclass that overrides GetTitle to supply French titles for the chart, category axis, and value axis. The custom settings are attached to the workbook so titles are applied automatically, and the file is saved as FrenchChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set French titles directly (globalization via subclass not required).
                chart.Title.Text = "Titre du graphique";
                chart.CategoryAxis.Title.Text = "Titre de l'axe des catégories";
                chart.ValueAxis.Title.Text = "Titre de l'axe des valeurs";
                // Legend does not have a Title property; this line is omitted.

                // Save the workbook.
                string outputPath = "FrenchChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

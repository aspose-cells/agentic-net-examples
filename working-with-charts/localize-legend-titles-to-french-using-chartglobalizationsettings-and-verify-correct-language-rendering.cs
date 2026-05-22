using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LocalizeLegendFrench
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Ventes Mensuelles";

        // Create SettableChartGlobalizationSettings with French legend terms
        SettableChartGlobalizationSettings frenchSettings = new SettableChartGlobalizationSettings();
        frenchSettings.SetLegendIncreaseName("Augmentation");
        frenchSettings.SetLegendDecreaseName("Diminution");
        frenchSettings.SetLegendTotalName("Total");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = frenchSettings
        };

        // Verify that the French terms are set correctly
        Console.WriteLine("Legend Increase (FR): " + frenchSettings.GetLegendIncreaseName());
        Console.WriteLine("Legend Decrease (FR): " + frenchSettings.GetLegendDecreaseName());
        Console.WriteLine("Legend Total (FR): " + frenchSettings.GetLegendTotalName());

        // Save the workbook
        workbook.Save("LocalizedLegendFrench.xlsx");
    }
}
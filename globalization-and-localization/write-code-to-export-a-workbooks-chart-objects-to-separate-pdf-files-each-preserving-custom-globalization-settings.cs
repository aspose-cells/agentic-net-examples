using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ExportChartsToPdf
{
    static void Main()
    {
        // Load the workbook containing charts
        Workbook workbook = new Workbook("input.xlsx");

        // Create custom chart globalization settings
        var chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetChartTitleName("Custom Chart Title");
        chartSettings.SetSeriesName("Custom Series");
        chartSettings.SetLegendIncreaseName("Increase");
        chartSettings.SetLegendDecreaseName("Decrease");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = chartSettings
        };

        // Iterate through all worksheets and their charts
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet sheet = workbook.Worksheets[wsIndex];

            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];

                // Generate a unique PDF file name for each chart
                string pdfFileName = $"{sheet.Name}_Chart{chartIndex + 1}.pdf";

                // Export the chart to a PDF file, preserving the custom globalization settings
                chart.ToPdf(pdfFileName);

                Console.WriteLine($"Exported chart {chartIndex + 1} from sheet '{sheet.Name}' to '{pdfFileName}'.");
            }
        }
    }
}
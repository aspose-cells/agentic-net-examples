using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    public static class ChartGlobalizationHelper
    {
        public static void CreateChartWithLanguage(string languageCode, string outputPath)
        {
            try
            {
                // Ensure the output directory exists.
                string dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare sample data for the chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Series");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Build language‑specific chart globalization settings.
                SettableChartGlobalizationSettings chartSettings = GetChartSettingsForLanguage(languageCode);

                // Apply globalization settings to the workbook.
                SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings
                {
                    ChartSettings = chartSettings
                };
                workbook.Settings.GlobalizationSettings = globalSettings;

                // Create a column chart.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set localized chart title.
                chart.Title.Text = chartSettings.GetChartTitleName();

                // Save the workbook.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating chart: {ex.Message}");
                throw;
            }
        }

        private static SettableChartGlobalizationSettings GetChartSettingsForLanguage(string languageCode)
        {
            var settings = new SettableChartGlobalizationSettings();

            switch (languageCode?.ToLowerInvariant())
            {
                case "en":
                    settings.SetChartTitleName("Sales Overview");
                    settings.SetSeriesName("Sales Series");
                    settings.SetLegendIncreaseName("Increase");
                    settings.SetLegendDecreaseName("Decrease");
                    settings.SetOtherName("Other");
                    break;
                case "fr":
                    settings.SetChartTitleName("Aperçu des ventes");
                    settings.SetSeriesName("Série des ventes");
                    settings.SetLegendIncreaseName("Augmentation");
                    settings.SetLegendDecreaseName("Diminution");
                    settings.SetOtherName("Autre");
                    break;
                case "de":
                    settings.SetChartTitleName("Verkaufsübersicht");
                    settings.SetSeriesName("Verkaufsreihe");
                    settings.SetLegendIncreaseName("Zunahme");
                    settings.SetLegendDecreaseName("Abnahme");
                    settings.SetOtherName("Andere");
                    break;
                case "es":
                    settings.SetChartTitleName("Resumen de ventas");
                    settings.SetSeriesName("Serie de ventas");
                    settings.SetLegendIncreaseName("Incremento");
                    settings.SetLegendDecreaseName("Disminución");
                    settings.SetOtherName("Otro");
                    break;
                default:
                    // Fallback to English.
                    settings.SetChartTitleName("Sales Overview");
                    settings.SetSeriesName("Sales Series");
                    settings.SetLegendIncreaseName("Increase");
                    settings.SetLegendDecreaseName("Decrease");
                    settings.SetOtherName("Other");
                    break;
            }

            return settings;
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Example usage: create a French chart.
                string language = "fr";
                string outputFile = "Chart_French.xlsx";

                ChartGlobalizationHelper.CreateChartWithLanguage(language, outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error in Main: {ex.Message}");
            }
        }
    }
}
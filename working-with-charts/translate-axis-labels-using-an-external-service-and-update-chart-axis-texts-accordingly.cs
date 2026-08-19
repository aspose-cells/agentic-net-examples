// Title: Translate Aspose.Cells Chart Axis Labels Using an External Service (C#/.NET)
// Description: Creates a workbook with a column chart, extracts the current X‑ and Y‑axis texts, translates each label via an async external service, writes the translated category names back to the sheet, reassigns the category range, applies a custom number format to the value axis, and saves the file as TranslatedChart.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart axis translation | localize chart labels | external translation API | async label translation | category axis update | value axis number format | Excel chart localization | Aspose.Cells Chart API
// Common Searches: Aspose.Cells translate chart axis labels | localize Excel chart axis text C# | update category axis after chart creation Aspose.Cells | apply custom number format to value axis Aspose.Cells | call translation service from Aspose.Cells example | async translation of chart labels .NET
// Developer Intent: Replace chart axis texts with translations obtained from an external service.
// Use Cases: Localize chart axis labels by calling a translation API for each label and writing the results back to the source cells. | Refresh the chart after updating the category data range so the new translated labels appear automatically. | Show numeric axis values with a language‑specific suffix using a custom number format.
// AI Prompts: Generate C# code that calls Azure Translator to translate Aspose.Cells chart axis labels and updates the chart. | Explain how to set a dynamic number format on the value axis to append a language code suffix to each numeric label. | Provide steps to recalculate and redraw an Aspose.Cells chart after modifying worksheet cells that supply category data.

using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisTranslation
{
    // Creates a workbook with a column chart, extracts the current X‑ and Y‑axis texts, translates each label via an async external service, writes the translated category names back to the sheet, reassigns the category range, applies a custom number format to the value axis, and saves the file as TranslatedChart.xlsx.
    class Program
    {
        // Mock external translation service – replace with real implementation as needed
        private static async Task<string> TranslateAsync(string text, string targetLanguage = "es")
        {
            // Simulate async call latency
            await Task.Delay(10);
            // Simple mock: append language code
            return $"{text}_{targetLanguage}";
        }

        static async Task Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Calculate the chart so that axis labels are generated
                chart.Calculate();

                // ---------- Retrieve current axis labels ----------
                // Value axis (Y‑axis) labels
                string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

                // Category axis (X‑axis) labels
                string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();

                // ---------- Translate labels ----------
                for (int i = 0; i < valueAxisLabels.Length; i++)
                {
                    valueAxisLabels[i] = await TranslateAsync(valueAxisLabels[i]);
                }

                for (int i = 0; i < categoryAxisLabels.Length; i++)
                {
                    categoryAxisLabels[i] = await TranslateAsync(categoryAxisLabels[i]);
                }

                // ---------- Update axis labels ----------
                // Write translated category names back to the worksheet.
                for (int i = 0; i < categoryAxisLabels.Length; i++)
                {
                    // Cells A2, A3, ... hold the original category names
                    sheet.Cells[i + 2, 0].PutValue(categoryAxisLabels[i]);
                }

                // Re‑assign the category data range to reflect the new texts
                chart.NSeries.CategoryData = "A2:A4";

                // For value‑axis labels (numeric), set a custom number format that includes the language suffix.
                // Example: 120 becomes 120_es
                chart.ValueAxis.TickLabels.NumberFormat = "\"_es\"0";

                // ---------- Save the workbook ----------
                workbook.Save("TranslatedChart.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

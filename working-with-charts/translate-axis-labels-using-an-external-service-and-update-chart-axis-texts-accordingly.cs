// Title: Translate Excel chart axis labels via an external service using Aspose.Cells for .NET
// Description: This example shows how to load or create a workbook, generate a column chart, read the category axis texts, send each label to a translation service, write the translated strings back to the source cells, recalculate the chart so the new labels appear, and save the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | chart axis translation | localize Excel chart | update chart axis programmatically | external translation API | recalculate chart | Excel automation | category axis labels
// Common Searches: Aspose.Cells change chart axis labels after creation | C# translate Excel chart axis text using API | Refresh Aspose.Cells chart after editing source cells | Localize Excel charts programmatically .NET | How to update category axis values in Aspose.Cells
// Developer Intent: Replace the original category axis labels of an Aspose.Cells chart with translated text retrieved from an external service and produce a localized workbook.
// Use Cases: Extract current axis labels, translate them via a web service, write the results back to the worksheet, and refresh the chart. | Generate multilingual Excel reports where chart captions are automatically localized without manual editing. | Integrate a translation API into an automated reporting pipeline to produce region‑specific visualizations.
// AI Prompts: Write C# code that reads chart axis texts with Aspose.Cells, calls a translation API, updates the source cells, and refreshes the chart. | Explain how to force Aspose.Cells to redraw a chart after modifying the underlying data range. | Provide error‑handling patterns for failed translation service calls while updating Excel chart labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to load or create a workbook, generate a column chart, read the category axis texts, send each label to a translation service, write the translated strings back to the source cells, recalculate the chart so the new labels appear, and save the file as XLSX.
    public class TranslateChartAxisLabels
    {
        // Dummy external translation service – replace with real implementation
        private static string TranslateLabel(string original)
        {
            // For demonstration we just prepend a marker.
            return $"[Translated] {original}";
        }

        public static void Run()
        {
            try
            {
                // ---------- Create / Load ----------
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample category labels in column A
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                // Sample values in column B
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories

                // ---------- Calculate chart to generate axis texts ----------
                chart.Calculate();

                // Get current category axis labels
                Axis categoryAxis = chart.CategoryAxis;
                string[] originalLabels = categoryAxis.GetAxisTexts();

                // Translate each label using the external service
                string[] translatedLabels = new string[originalLabels.Length];
                for (int i = 0; i < originalLabels.Length; i++)
                {
                    translatedLabels[i] = TranslateLabel(originalLabels[i]);
                }

                // Write translated labels back to the source cells (A2, A3, ...)
                // Assuming the category data range starts at row 2 in column A
                int startRow = 2; // Excel rows are 1‑based; row 2 corresponds to "A2"
                for (int i = 0; i < translatedLabels.Length; i++)
                {
                    int row = startRow + i;
                    sheet.Cells[$"A{row}"].PutValue(translatedLabels[i]);
                }

                // Re‑calculate the chart so it picks up the updated cell values
                chart.Calculate();

                // ---------- Save ----------
                string outputPath = "TranslatedAxisLabelsChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TranslateChartAxisLabels.Run();
        }
    }
}

// Title: Translate Excel chart category axis labels with Aspose.Cells using an external REST API in C#
// AI Prompts: Generate C# code that extracts the category axis texts from an Aspose.Cells chart, calls a translation REST endpoint for each label, and returns the translated strings. | Add comprehensive error handling so that if a translation request fails, the original label is kept, then write the final labels back to the worksheet cells. | Demonstrate how to recalculate the chart after updating the source cells and save the workbook as an XLSX file with Aspose.Cells.
// Common Searches: how to use Aspose.Cells to translate chart axis labels via a web service in C# | c# Aspose.Cells update category axis text after calling external API | refresh Excel chart in Aspose.Cells after programmatically changing cell values | handle translation API failures when modifying Excel chart labels with Aspose.Cells
// Tags: Aspose.Cells chart axis localization via REST | C# update category axis text in Excel chart | recalculate Aspose.Cells chart after cell edit | error handling external translation service Aspose.Cells | save workbook as XLSX using Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a column chart, read its category axis labels, translate each label through an async HTTP call, write the translated labels back to the source cells, recalculate the chart to reflect the changes, and save the result as an XLSX file.
class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // Create workbook (lifecycle rule placeholder)
            Workbook workbook = CreateWorkbook();

            Worksheet ws = workbook.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(8000);
            ws.Cells["B3"].PutValue(4000);
            ws.Cells["B4"].PutValue(-8000);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Calculate chart to generate axis texts
            chart.Calculate();

            // Retrieve current axis labels
            string[] originalLabels = chart.CategoryAxis.GetAxisTexts();

            // Translate each label via an external service (fallback to original on failure)
            string[] translatedLabels = new string[originalLabels.Length];
            for (int i = 0; i < originalLabels.Length; i++)
            {
                try
                {
                    translatedLabels[i] = await TranslateLabelAsync(originalLabels[i]);
                }
                catch (Exception ex)
                {
                    // Log the error and keep the original label
                    Console.WriteLine($"Translation failed for \"{originalLabels[i]}\": {ex.Message}");
                    translatedLabels[i] = originalLabels[i];
                }
            }

            // Write translated labels back to the source cells (A2, A3, ...)
            for (int i = 0; i < translatedLabels.Length; i++)
            {
                ws.Cells[i + 1, 0].PutValue(translatedLabels[i]); // Row i+1, column 0 = "A"
            }

            // Recalculate chart to reflect updated labels
            chart.Calculate();

            // Save workbook (lifecycle rule placeholder)
            SaveWorkbook(workbook, "TranslatedChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Calls an external translation service (mock implementation with graceful fallback)
    static async Task<string> TranslateLabelAsync(string text)
    {
        // Replace with a real endpoint if available
        const string endpoint = "https://api.example.com/translate";

        using (HttpClient client = new HttpClient())
        {
            // Build request URL
            string requestUrl = $"{endpoint}?text={Uri.EscapeDataString(text)}&to=es";

            // Send request
            HttpResponseMessage response = await client.GetAsync(requestUrl);
            response.EnsureSuccessStatusCode();

            // Return translated text
            return await response.Content.ReadAsStringAsync();
        }
    }

    // Placeholder for the create lifecycle rule
    static Workbook CreateWorkbook()
    {
        // The actual implementation is supplied by the rule engine
        return new Workbook();
    }

    // Placeholder for the save lifecycle rule
    static void SaveWorkbook(Workbook wb, string path)
    {
        try
        {
            // Ensure the directory exists
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            wb.Save(path, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to \"{path}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}

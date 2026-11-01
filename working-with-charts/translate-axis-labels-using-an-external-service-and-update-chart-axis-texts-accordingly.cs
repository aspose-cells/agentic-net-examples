// Title: Translate Excel chart axis titles to another language using Aspose.Cells and an external translation API (C#)
// AI Prompts: Write C# code that opens an XLSX workbook with Aspose.Cells, iterates all charts, calls a web API to translate each axis title, and replaces the chart titles with the translated text. | Adapt the example to use a different translation endpoint and change the target language parameter to French while preserving async error handling and fallback behavior. | Add logging that records the original axis title, the requested language, and the translated result for every chart processed.
// Common Searches: how to translate Excel chart axis labels using Aspose.Cells in C# | c# Aspose.Cells update chart axis title with external translation service | iterate through all charts in a workbook and change axis text via web API | call REST translation API from Aspose.Cells chart processing code | save workbook after modifying chart axis titles in .NET
// Tags: Aspose.Cells translate chart axis titles | C# call external translation API | update Excel chart axis text programmatically | async web request for chart label localization | save modified workbook with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

// The sample loads an XLSX workbook with Aspose.Cells, loops through every worksheet and chart, translates the Category (X) and Value (Y) axis titles by invoking an external translation service, replaces the original titles with the translated strings, and saves the updated workbook, handling missing files and translation errors gracefully.
class Program
{
    static async Task Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Translate Category (X) axis title if it exists
                    Axis categoryAxis = chart.CategoryAxis;
                    if (categoryAxis?.Title != null && !string.IsNullOrEmpty(categoryAxis.Title.Text))
                    {
                        string translatedTitle = await TranslateTextAsync(categoryAxis.Title.Text);
                        categoryAxis.Title.Text = translatedTitle;
                    }

                    // Translate Value (Y) axis title if it exists
                    Axis valueAxis = chart.ValueAxis;
                    if (valueAxis?.Title != null && !string.IsNullOrEmpty(valueAxis.Title.Text))
                    {
                        string translatedTitle = await TranslateTextAsync(valueAxis.Title.Text);
                        valueAxis.Title.Text = translatedTitle;
                    }
                }
            }

            // Save workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Calls an external translation service and returns the translated text
    static async Task<string> TranslateTextAsync(string text)
    {
        const string translationServiceUrl = "https://api.example.com/translate";

        try
        {
            using (HttpClient client = new HttpClient())
            {
                var payload = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("text", text),
                    new KeyValuePair<string, string>("targetLang", "es")
                });

                HttpResponseMessage response = await client.PostAsync(translationServiceUrl, payload);
                response.EnsureSuccessStatusCode();

                // Assuming the service returns the translated text as plain string
                return await response.Content.ReadAsStringAsync();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Translation failed: {ex.Message}");
            // Return original text if translation fails
            return text;
        }
    }
}

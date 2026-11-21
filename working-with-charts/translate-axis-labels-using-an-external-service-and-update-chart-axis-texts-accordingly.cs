// Title: Translate Excel chart axis titles with an external async service using Aspose.Cells for .NET
// Common Searches: how to use Aspose.Cells to translate chart axis titles with a web API in C# | C# example for updating Excel chart X and Y axis labels via external translation service | async call to translate Excel chart axis text before saving workbook using Aspose.Cells | replace Excel chart axis titles with translated strings using Aspose.Cells .NET | load workbook, iterate charts, translate axis titles using async method Aspose.Cells
// Tags: Aspose.Cells chart axis localization | C# async translation integration for Excel charts | programmatic update of Excel chart axis text | external API call to modify chart titles Aspose.Cells | Excel workbook chart label translation .NET

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, walks through each worksheet and its charts, and asynchronously translates the Category (X) and Value (Y) axis titles using a placeholder TranslateAsync method that represents an external translation service. The translated strings replace the original axis titles, and the workbook is saved to a new file. Includes error handling for missing input files and demonstrates async API integration with Aspose.Cells in C#.
class Program
{
    // Mock translation method – replace with actual external service call
    static async Task<string> TranslateAsync(string text)
    {
        // Example: call an external API here and return the translated text
        // For demonstration, append "_translated"
        return await Task.FromResult(text + "_translated");
    }

    static async Task Main(string[] args)
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Translate Category (X) Axis title if it exists
                    Axis categoryAxis = chart.CategoryAxis;
                    if (categoryAxis.Title != null && !string.IsNullOrEmpty(categoryAxis.Title.Text))
                    {
                        string translated = await TranslateAsync(categoryAxis.Title.Text);
                        categoryAxis.Title.Text = translated;
                    }

                    // Translate Value (Y) Axis title if it exists
                    Axis valueAxis = chart.ValueAxis;
                    if (valueAxis.Title != null && !string.IsNullOrEmpty(valueAxis.Title.Text))
                    {
                        string translated = await TranslateAsync(valueAxis.Title.Text);
                        valueAxis.Title.Text = translated;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

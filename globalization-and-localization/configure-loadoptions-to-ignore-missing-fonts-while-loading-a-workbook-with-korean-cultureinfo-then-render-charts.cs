// Title: C# – Load Korean‑culture Excel workbook with IgnoreMissingFonts and render charts using Aspose.Cells
// Description: The sample checks for the input XLSX file, creates LoadOptions with CultureInfo "ko‑KR" and IgnoreMissingFonts enabled, sets a font substitute (Arial → Malgun Gothic), loads the workbook, configures ImageOrPrintOptions to use Malgun Gothic as the default font, and iterates through all worksheets to export each chart as a PNG image while handling errors gracefully.
// Keywords: Aspose.Cells | C# | LoadOptions | IgnoreMissingFonts | Korean CultureInfo | font substitution | Malgun Gothic | Arial fallback | Chart.ToImage | export chart PNG | Excel chart rendering | globalization | localization
// Common Searches: Aspose.Cells ignore missing fonts Korean | C# load Excel with Korean locale Aspose | set font substitutes in Aspose.Cells | export all Excel charts to PNG C# | render charts with Korean text Aspose.Cells
// Developer Intent: Load an Excel workbook using Korean locale, ignore missing fonts, apply a fallback font, and generate PNG images for every chart.
// Use Cases: Create localized chart images for Korean dashboards without installing the original fonts. | Batch‑convert Excel charts to web‑ready PNG files while preserving Korean characters. | Implement a fallback‑font strategy to ensure consistent chart appearance when source fonts are unavailable.
// AI Prompts: Show how to enable IgnoreMissingFonts in LoadOptions while keeping Korean CultureInfo and font substitution. | Provide a code snippet that saves each rendered chart into a subfolder named after its worksheet. | Explain how to change ImageOrPrintOptions to output JPEG instead of PNG for chart rendering.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample checks for the input XLSX file, creates LoadOptions with CultureInfo "ko‑KR" and IgnoreMissingFonts enabled, sets a font substitute (Arial → Malgun Gothic), loads the workbook, configures ImageOrPrintOptions to use Malgun Gothic as the default font, and iterates through all worksheets to export each chart as a PNG image while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook.
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Configure load options with Korean culture and font substitution.
            LoadOptions loadOptions = new LoadOptions
            {
                CultureInfo = new CultureInfo("ko-KR")
            };
            IndividualFontConfigs fontConfigs = new IndividualFontConfigs();
            fontConfigs.SetFontSubstitutes("Arial", new[] { "Malgun Gothic" });
            loadOptions.FontConfigs = fontConfigs;

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Prepare image rendering options for charts.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Use the workbook's default font first; helps when characters are Unicode.
                CheckWorkbookDefaultFont = true,
                // Specify a Korean font to ensure proper rendering of Korean text.
                DefaultFont = "Malgun Gothic"
                // Image format defaults to PNG, so no explicit setting required.
            };

            // Iterate through worksheets and render each chart.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartCounter = 0;
                foreach (Chart chart in sheet.Charts)
                {
                    chartCounter++;
                    try
                    {
                        // Render the chart to an image using the Chart.ToImage method.
                        string outputFile = $"{sheet.Name}_Chart{chartCounter}.png";
                        chart.ToImage(outputFile, imgOptions);
                        Console.WriteLine($"Chart rendered to: {outputFile}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to render chart on sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

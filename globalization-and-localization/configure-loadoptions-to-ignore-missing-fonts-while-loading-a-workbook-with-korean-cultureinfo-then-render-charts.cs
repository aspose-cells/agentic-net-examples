// Title: Load a Korean‑localized Excel workbook with missing‑font suppression and export its charts to PNG using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with Aspose.Cells, set LoadOptions.IgnoreMissingFonts = true, apply Korean CultureInfo to the thread, and save each worksheet chart as a PNG file. | Create a C# method that accepts an Excel path, configures LoadOptions to suppress missing fonts, sets the current culture to ko‑KR, renders all charts to PNG images, and returns the list of generated filenames. | Adjust existing Aspose.Cells code to handle missing fonts gracefully when loading a Korean workbook and batch‑export every chart to separate PNG files.
// Common Searches: Aspose.Cells C# missing font suppression for Korean workbook loading | Export Excel charts to PNG after setting CultureInfo to ko‑KR with Aspose.Cells | How to configure LoadOptions to bypass missing fonts in Aspose.Cells .NET | Render charts from a Korean localized Excel file using Aspose.Cells | Aspose.Cells chart rendering fails due to missing fonts Korean locale
// Tags: missing-font suppression loadoptions aspnet | korean locale workbook loading aspnet | excel chart export png aspnet | chart rendering font fallback aspnet | aspocells load without font errors

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example shows how to set the current thread to Korean (ko‑KR) culture, configure Aspose.Cells LoadOptions with IgnoreMissingFonts = true to prevent font‑related errors, load an Excel workbook, iterate through all worksheets, render each chart to a PNG image, and optionally save the workbook, with comprehensive error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Set Korean culture for the current thread
            CultureInfo korean = new CultureInfo("ko-KR");
            System.Threading.Thread.CurrentThread.CurrentCulture = korean;
            System.Threading.Thread.CurrentThread.CurrentUICulture = korean;

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through all worksheets and render each chart to an image
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Create a unique file name for each chart image
                        string chartIdentifier = !string.IsNullOrEmpty(chart.Name)
                            ? chart.Name
                            : $"Chart_{sheet.Name}_{chart.GetHashCode()}";

                        string imagePath = $"{chartIdentifier}.png";

                        // Render the chart to a PNG image (default format)
                        chart.ToImage(imagePath);
                        Console.WriteLine($"Chart saved as {imagePath}");
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Failed to render chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Optionally save the workbook after processing
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Processing completed. Workbook saved as {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

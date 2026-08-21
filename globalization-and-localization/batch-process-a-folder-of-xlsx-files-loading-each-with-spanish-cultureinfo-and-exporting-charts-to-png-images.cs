// Title: C# – Batch export Excel charts to PNG with Spanish (es‑ES) locale using Aspose.Cells
// Description: A console utility that scans a folder for *.xlsx files, loads each workbook with a Spanish (es‑ES) CultureInfo via LoadOptions, iterates every worksheet and chart, and saves each chart as an individual PNG file in a target directory. Includes robust folder validation and per‑chart error handling.
// Keywords: Aspose.Cells | C# | export chart to PNG | batch process Excel files | Spanish locale | es-ES CultureInfo | LoadOptions | chart extraction | Excel automation | globalization | localization | folder scanning | chart image generation | Aspose.Cells for .NET
// Common Searches: Aspose.Cells batch export charts to PNG | load Excel workbook with Spanish culture C# | export all charts from multiple XLSX files | C# code to convert Excel charts to images | Aspose.Cells chart image generation with locale
// Developer Intent: Automatically generate PNG images for every chart in each XLSX workbook within a directory, applying the Spanish (es‑ES) culture during load.
// Use Cases: Create locale‑specific chart graphics for a Spanish‑language reporting portal. | Produce a library of PNG assets for marketing collateral from a collection of Excel workbooks. | Validate that chart rendering respects Spanish number formats and date conventions before publishing.
// AI Prompts: Write C# code that uses Aspose.Cells to batch export charts from XLSX files with a specified CultureInfo and custom file naming. | Explain best practices for handling exceptions when exporting charts to PNG inside a loop with Aspose.Cells. | Show how to adapt the sample to output SVG files while keeping the Spanish locale settings.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A console utility that scans a folder for *.xlsx files, loads each workbook with a Spanish (es‑ES) CultureInfo via LoadOptions, iterates every worksheet and chart, and saves each chart as an individual PNG file in a target directory. Includes robust folder validation and per‑chart error handling.
class ExportChartsBatch
{
    static void Main(string[] args)
    {
        // Input folder containing XLSX files
        string inputFolder = @"C:\InputExcelFiles";
        // Output folder for PNG images
        string outputFolder = @"C:\ExportedCharts";

        try
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create output directory: {ex.Message}");
            return;
        }

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Process each XLSX file in the folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // LoadOptions with Spanish culture (es-ES)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CultureInfo = new CultureInfo("es-ES")
                };

                // Load the workbook
                using (Workbook workbook = new Workbook(filePath, loadOptions))
                {
                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through all charts in the worksheet
                        foreach (Chart chart in sheet.Charts)
                        {
                            // Determine chart index within the worksheet
                            int chartIdx = sheet.Charts.IndexOf(chart);

                            // Build a unique file name for each chart
                            string chartName = string.IsNullOrEmpty(chart.Name) ? $"Chart_{chartIdx}" : chart.Name;
                            string baseFileName = Path.GetFileNameWithoutExtension(filePath);
                            string imageFileName = $"{baseFileName}_{chartName}.png";
                            string imagePath = Path.Combine(outputFolder, imageFileName);

                            try
                            {
                                // Export the chart to PNG
                                chart.ToImage(imagePath);
                            }
                            catch (Exception imgEx)
                            {
                                Console.WriteLine($"Failed to export chart '{chartName}' from '{filePath}': {imgEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Chart export completed.");
    }
}

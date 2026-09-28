// Title: Batch export all charts from multiple XLSX workbooks to localized PNG files using Chinese culture settings with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory for .xlsx files, loads each workbook with CultureInfo set to zh-CN, and saves every chart as a PNG using Aspose.Cells. | Create a method that configures ImageOrPrintOptions for PNG output and applies it to each chart while handling exceptions during bulk export. | Generate code that builds output filenames containing the source workbook name, worksheet index, and chart index when exporting charts with Aspose.Cells.
// Common Searches: how to export all charts from multiple Excel files to PNG with Aspose.Cells and Chinese locale | C# batch process XLSX workbooks to generate chart images using Aspose.Cells | set workbook cultureinfo to zh-CN for chart rendering in Aspose.Cells | save Excel chart as PNG with sheet and chart index in filename using Aspose.Cells | console app to convert Excel charts to localized PNG images in bulk
// Tags: batch chart export PNG Aspose.Cells C# | workbook cultureinfo zh-CN Aspose.Cells | chart image localization Chinese Aspose.Cells | output filename includes workbook sheet chart indices | console application bulk Excel chart conversion

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// A C# console utility scans a given folder for .xlsx files, sets each workbook's CultureInfo to zh-CN, iterates through all worksheets and charts, and exports each chart as a PNG file named with the workbook, sheet, and chart indices, using Aspose.Cells' ImageOrPrintOptions.
class ChartExporter
{
    static void Main()
    {
        // Input and output directories
        string inputFolder = @"C:\Input";
        string outputFolder = @"C:\Output";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load workbook
                Workbook workbook = new Workbook(filePath);
                workbook.Settings.CultureInfo = new CultureInfo("zh-CN");

                int chartCounter = 0;
                string baseFileName = Path.GetFileNameWithoutExtension(filePath);

                // Iterate through worksheets and their charts
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            // Configure image export options
                            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                            {
                                // Specify image type (PNG)
                                ImageType = ImageType.Png
                            };

                            // Build output file name
                            string chartFileName = $"{baseFileName}_Sheet{sheet.Index}_Chart{chartCounter}.png";
                            string outputPath = Path.Combine(outputFolder, chartFileName);

                            // Export chart to file (outputPath first, then options)
                            chart.ToImage(outputPath, imgOptions);

                            chartCounter++;
                        }
                        catch (Exception exChart)
                        {
                            Console.WriteLine($"Error exporting chart in file '{filePath}': {exChart.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}

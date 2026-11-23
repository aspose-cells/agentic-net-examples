// Title: Load an Excel workbook with Korean (ko‑KR) CultureInfo while ignoring missing fonts and export all charts to transparent PNG images using Aspose.Cells for .NET
// Common Searches: c# aspnet cells load workbook with Korean culture and ignore missing fonts | how to export all Excel charts to PNG with transparent background using Aspose.Cells | aspocells set LoadOptions.IgnoreMissingFonts example | render charts from a workbook loaded with specific CultureInfo in Aspose.Cells | batch save each chart as a separate image file Aspose.Cells C#
// Tags: loadoptions ignoremissingfonts cultureinfo ko-kr | aspocells render chart to png transparent | export excel charts as images c# aspnet | set workbook culture for localization aspocells | chart toimage batch processing aspocells

using System;
using System.Globalization;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

// The sample creates LoadOptions with Korean (ko‑KR) CultureInfo and the IgnoreMissingFonts flag, loads an Excel workbook, iterates through all worksheets and their charts, and renders each chart to a uniquely named PNG file with a transparent background using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Configure LoadOptions with Korean culture
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("ko-KR")
            };

            // Load the workbook with the specified options
            var workbook = new Workbook(inputPath, loadOptions);

            // Render each chart in the workbook to an image file
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Set image rendering options
                        var imgOptions = new ImageOrPrintOptions
                        {
                            // PNG is the default format; Transparent background is enabled
                            Transparent = true
                        };

                        // Build a unique file name for the chart image
                        string chartIdentifier = !string.IsNullOrEmpty(chart.Name)
                            ? chart.Name
                            : Guid.NewGuid().ToString("N");
                        string chartFileName = $"Chart_{sheet.Name}_{chartIdentifier}.png";

                        // Render the chart to the image file (file name first, then options)
                        chart.ToImage(chartFileName, imgOptions);
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Failed to render chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Optionally save the workbook after processing
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

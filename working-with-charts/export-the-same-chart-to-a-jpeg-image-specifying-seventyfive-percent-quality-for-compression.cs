// Title: Export a worksheet chart to a JPEG file with 75% compression using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, selects the first chart on the first worksheet, and saves it as a JPEG image with 75% quality using Aspose.Cells. | Show how to configure ImageOrPrintOptions to set JPEG compression quality when converting a chart to an image with Aspose.Cells. | Write a C# routine that iterates over all charts in a worksheet and exports each one to a separate JPEG file with 75% compression using Aspose.Cells.
// Common Searches: Aspose.Cells C# export chart as JPEG with specific quality | how to set JPEG quality when saving Excel chart using Aspose.Cells | C# convert Excel chart to JPEG with 75% compression | ImageOrPrintOptions Quality property example for chart image export Aspose.Cells | export multiple charts from worksheet to JPEG files using Aspose.Cells .NET
// Tags: chart to JPEG export Aspose.Cells | ImageOrPrintOptions JPEG quality setting | export Excel chart as image C# | Aspose.Cells chart image conversion with compression | batch export worksheet charts to JPEG Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads input.xlsx, checks for a chart on the first worksheet, and uses ImageOrPrintOptions with Quality=75 to export that chart as chart.jpg in JPEG format.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "chart.jpg";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Configure image export options (JPEG, 75% quality)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Quality is applicable for JPEG format
                Quality = 75
            };

            // Export the chart directly to an image file
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine($"Chart exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

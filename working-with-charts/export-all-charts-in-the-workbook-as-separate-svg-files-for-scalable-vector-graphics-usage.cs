// Title: Export every chart from an Excel workbook to individual SVG files using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, iterates over all worksheets and their charts, and saves each chart as a separate SVG file with Aspose.Cells. | Show how to configure ImageOrPrintOptions for SVG output when calling Chart.ToImage in Aspose.Cells .NET. | Create a reusable method that returns the file paths of SVG images generated from all charts in a specified Excel file using Aspose.Cells.
// Common Searches: C# Aspose.Cells export each worksheet chart to SVG | How to save Excel charts as separate SVG files with Aspose.Cells .NET | Iterate through workbook charts and generate SVG images using Aspose.Cells | Aspose.Cells Chart.ToImage example for SVG output in C#
// Tags: Aspose.Cells chart export to SVG | C# iterate workbook charts Aspose.Cells | ImageOrPrintOptions SVG configuration Aspose.Cells | multiple chart SVG files generation Aspose.Cells | export Excel charts as scalable vector graphics .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, loops through each worksheet and its charts, and uses Aspose.Cells' ImageOrPrintOptions with the Chart.ToImage method to export every chart as an individual SVG file named by sheet and chart indices, while handling missing files and export errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            int sheetIndex = 0;
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIndex = 0;
                // Iterate through each chart on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Configure image options (format inferred from file extension)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                        // Build a unique file name for each chart
                        string svgFileName = $"Chart_Sheet{sheetIndex}_Chart{chartIndex}.svg";

                        // Export the chart as an SVG file
                        chart.ToImage(svgFileName, imgOptions);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart {chartIndex} on sheet {sheetIndex}: {ex.Message}");
                    }

                    chartIndex++;
                }
                sheetIndex++;
            }

            Console.WriteLine("Chart export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

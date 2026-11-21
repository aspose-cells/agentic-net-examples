// Title: Export every chart in an Excel workbook to high‑resolution PNG files with Aspose.Cells for .NET
// Common Searches: C# code to loop through all worksheets and save each chart as a PNG using Aspose.Cells | How to set 300 DPI when exporting Excel charts to images with Aspose.Cells .NET | Batch export charts from an .xlsx file to PNG files in a .NET application | Aspose.Cells example for exporting multiple charts with custom image options | Save Excel chart images with sheet name and index using Aspose.Cells C#
// Tags: Aspose.Cells chart to PNG conversion C# | high DPI chart image generation Aspose.Cells | iterate worksheets for chart extraction Aspose.Cells | ImageOrPrintOptions resolution setting Aspose.Cells | batch chart image creation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, iterates over each worksheet and its chart collection, configures ImageOrPrintOptions to 300 dpi, and exports every chart to a PNG file named with the worksheet and chart index, while handling missing files and export errors.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ChartCollection charts = sheet.Charts;

                // Process each chart
                for (int i = 0; i < charts.Count; i++)
                {
                    try
                    {
                        Chart chart = charts[i];

                        // Set image export options (default format is PNG)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            HorizontalResolution = 300,
                            VerticalResolution = 300
                        };

                        // Export the chart as an image file
                        string imageFileName = $"{sheet.Name}_Chart{i + 1}.png";
                        chart.ToImage(imageFileName, imgOptions);
                        Console.WriteLine($"Exported chart to '{imageFileName}'.");
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error exporting chart {i + 1} on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

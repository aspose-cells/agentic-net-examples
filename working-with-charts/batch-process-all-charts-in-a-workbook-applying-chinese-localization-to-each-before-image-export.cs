// Title: Batch apply Chinese font to all charts in an Excel workbook and export each chart as PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, iterates over every worksheet and chart, changes the chart title and legend fonts to a Chinese typeface (e.g., Microsoft YaHei), and saves each chart as a PNG file whose name includes the sheet index and chart identifier. | Create a naming convention for chart image files that combines the worksheet number and chart name, then adjust the code to export the charts with ImageOrPrintOptions and finally persist the workbook.
// Common Searches: how to change chart title font to Chinese in Aspose.Cells C# | export all charts from an Excel file to PNG using Aspose.Cells .NET | batch process workbook charts to set Microsoft YaHei font with Aspose.Cells | C# loop through worksheets and charts to localize fonts in Aspose.Cells | save each chart as separate image file naming by sheet index Aspose.Cells
// Tags: set chart title font Chinese Aspose.Cells | export workbook charts to PNG C# | batch process charts Aspose.Cells .NET | localize chart legend font Microsoft YaHei | image export options Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering; // Required for ImageOrPrintOptions

// The example loads an Excel workbook, iterates through every worksheet and its charts, applies a Chinese font (Microsoft YaHei) to chart titles and legends for proper localization, exports each chart as a PNG image with a filename that includes the sheet index and chart name, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartCounter = 0; // Counter for unnamed charts

                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Set a Chinese font for titles and legends to ensure proper rendering
                        if (chart.Title != null && chart.Title.IsVisible)
                        {
                            chart.Title.Font.Name = "Microsoft YaHei";
                        }

                        if (chart.Legend != null)
                        {
                            chart.Legend.Font.Name = "Microsoft YaHei";
                        }

                        // Prepare image export options (default format is PNG)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                        // Determine a safe chart name for the output file
                        chartCounter++;
                        string safeChartName = string.IsNullOrWhiteSpace(chart.Name)
                            ? $"Chart{chartCounter}"
                            : chart.Name;

                        string imagePath = $"Chart_Sheet{sheet.Index}_{safeChartName}.png";

                        // Export the chart to an image file
                        chart.ToImage(imagePath, imgOptions);

                        Console.WriteLine($"Chart exported to {imagePath}");
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Save the workbook (optional, as no structural changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

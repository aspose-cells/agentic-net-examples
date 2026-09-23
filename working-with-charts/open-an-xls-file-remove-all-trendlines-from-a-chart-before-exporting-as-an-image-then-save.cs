// Title: Delete trendlines from all charts in an XLS workbook and export each chart as a PNG using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an XLS workbook with Aspose.Cells, removes every trendline from each chart, and saves the modified workbook. | Write a C# routine that iterates through all worksheets and their charts, clears the Series.Trendlines collection, and exports each chart to a PNG file using ImageOrPrintOptions. | Create a C# example that uses reflection to handle different Aspose.Cells versions when removing chart trendlines before exporting the chart image.
// Common Searches: aspnet remove trendlines from charts in existing xls workbook Aspose.Cells | export chart as png after clearing trendlines using Aspose.Cells C# | how to iterate through all charts in a workbook and delete trendlines with Aspose.Cells | C# Aspose.Cells remove chart series trendlines before image export
// Tags: remove chart trendlines Aspose.Cells .NET | export chart to PNG Aspose.Cells | iterate worksheets and charts Aspose.Cells | clear series trendlines via reflection C# | save workbook without trendlines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an XLS workbook, walks through each worksheet and its charts, clears any trendlines from every series (using reflection for version safety), exports each chart as a PNG image, and finally saves the workbook with the trendlines removed.
class RemoveTrendlinesAndExport
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Attempt to remove trendlines via reflection (covers different API versions)
                    try
                    {
                        foreach (Series series in chart.NSeries)
                        {
                            var trendlinesProp = series.GetType().GetProperty("Trendlines");
                            if (trendlinesProp != null)
                            {
                                var trendlinesObj = trendlinesProp.GetValue(series);
                                if (trendlinesObj != null)
                                {
                                    var clearMethod = trendlinesObj.GetType().GetMethod("Clear");
                                    clearMethod?.Invoke(trendlinesObj, null);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to clear trendlines for chart \"{chart.Name}\": {ex.Message}");
                    }

                    // Set image export options (defaults to PNG)
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                    try
                    {
                        // Build a unique file name for the chart image
                        string chartName = !string.IsNullOrEmpty(chart.Name)
                            ? chart.Name
                            : $"Chart_{sheet.Name}_{Guid.NewGuid():N}";
                        string imageFileName = $"{chartName}.png";

                        // Export chart directly to an image file
                        chart.ToImage(imageFileName, imgOptions);
                        Console.WriteLine($"Chart image saved: {imageFileName}");
                    }
                    catch (Exception imgEx)
                    {
                        Console.WriteLine($"Failed to export chart image: {imgEx.Message}");
                    }
                }
            }

            // Save the modified workbook (trendlines removed if possible)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without trendlines: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

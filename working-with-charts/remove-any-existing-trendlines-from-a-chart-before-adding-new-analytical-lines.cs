// Title: How to remove all trendlines from a chart and then add a linear trendline using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that iterates through each series of the first worksheet chart, clears any existing trendlines, and then adds a linear trendline to the first series with its equation and R‑squared displayed. | Provide a version‑compatible C# snippet that uses reflection to access the Trendlines collection for each series, removes them, and creates a named linear trendline with display options enabled.
// Common Searches: aspocells c# remove existing trendlines before adding new trendline to chart | clear chart series trendlines Aspose.Cells example | add linear trendline with equation and R squared after deleting old trendlines in Aspose.Cells | how to use reflection to delete trendlines in older Aspose.Cells versions | c# Aspose.Cells chart trendline management tutorial
// Tags: Aspose.Cells remove chart trendlines C# | add linear trendline Aspose.Cells | trendline collection reflection Aspose.Cells | chart analytical lines management Aspose.Cells | Excel workbook chart modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, Series, Trendline, TrendlineType

namespace AsposeCellsTrendlineRemoval
{
    // The sample loads an Excel workbook, locates the first chart on the first worksheet, uses reflection to access each series' Trendlines collection, clears any existing trendlines, then adds a linear trendline to the first series with its name, equation, and R‑squared displayed, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load the workbook (use a default file name)
                string inputPath = "input.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // create a new workbook if the file is missing
                    Console.WriteLine($"Input file '{inputPath}' not found. Created a new workbook.");
                }

                // Assume the chart is on the first worksheet and is the first chart
                Worksheet sheet = workbook.Worksheets[0];

                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                }
                else
                {
                    try
                    {
                        Chart chart = sheet.Charts[0];

                        // Remove all existing trendlines from every series in the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Use reflection to keep compatibility with older library versions
                            var trendlinesProp = series.GetType().GetProperty("Trendlines");
                            if (trendlinesProp != null)
                            {
                                var trendlines = trendlinesProp.GetValue(series) as TrendlineCollection;
                                trendlines?.Clear();
                            }
                        }

                        // Add a linear trendline to the first series if present and API is supported
                        if (chart.NSeries.Count > 0)
                        {
                            Series firstSeries = chart.NSeries[0];
                            var trendlinesProp = firstSeries.GetType().GetProperty("Trendlines");
                            if (trendlinesProp != null)
                            {
                                var trendlines = trendlinesProp.GetValue(firstSeries) as TrendlineCollection;
                                if (trendlines != null)
                                {
                                    // Add returns the index of the new trendline; retrieve it to set properties
                                    int trendlineIndex = trendlines.Add(TrendlineType.Linear);
                                    Trendline linearTrendline = trendlines[trendlineIndex];
                                    linearTrendline.Name = "Linear Trendline";
                                    linearTrendline.DisplayEquation = true;
                                    linearTrendline.DisplayRSquared = true;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing chart: {ex.Message}");
                    }
                }

                // Save the workbook (use a default output file name)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

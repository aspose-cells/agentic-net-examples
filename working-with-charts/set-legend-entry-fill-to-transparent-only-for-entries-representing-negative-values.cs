// Title: How to make legend entries transparent for series with negative values in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code that loops through each series in an Aspose.Cells column chart, checks for any data points below zero, and sets the series' LegendEntry.BackgroundMode to Transparent. | Show an example of using Aspose.Range to read the values of a chart series and apply conditional formatting to the legend entry background when negative values are present.
// Common Searches: aspocells column chart legend entry transparent when series has negative values | c# detect negative data in chart series and change legend background Aspose.Cells | set legend entry background mode to transparent for negative values Aspose.Cells chart | conditional legend formatting based on series values Aspose.Cells .NET | how to read chart series range values in Aspose.Cells C#
// Tags: legend entry background transparent Aspose.Cells | negative data point detection chart series C# | conditional legend entry formatting Aspose.Cells | column chart legend background mode Aspose.Cells | read series data range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsLegendNegativeDemo
{
    // The example creates a workbook with two data series, adds a column chart, scans each series for values less than zero, and sets the corresponding legend entry's background mode to Transparent for any series that contain negative values before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data: first series contains negative values, second series only positive
                // Category labels
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1 (may have negative values)
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(-30); // negative value
                sheet.Cells["B4"].PutValue(70);

                // Series 2 (all positive)
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(60);
                sheet.Cells["C3"].PutValue(80);
                sheet.Cells["C4"].PutValue(90);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A4";

                // Iterate over each series and set legend entry fill to transparent
                // only if the series contains at least one negative data point
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    bool hasNegative = false;

                    // Series.Values holds the range address as a string (e.g., "B2:B4")
                    string valuesRange = series.Values as string;
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        try
                        {
                            // Get the range values as a 2‑D object array
                            AsposeRange range = sheet.Cells.CreateRange(valuesRange);
                            object[,] vals = range.Value as object[,];
                            if (vals != null)
                            {
                                foreach (object val in vals)
                                {
                                    if (val != null && double.TryParse(val.ToString(), out double d) && d < 0)
                                    {
                                        hasNegative = true;
                                        break;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to read range '{valuesRange}': {ex.Message}");
                        }
                    }

                    // If negative values are present, make the legend entry background transparent
                    if (hasNegative)
                    {
                        LegendEntry legendEntry = series.LegendEntry;
                        legendEntry.BackgroundMode = BackgroundMode.Transparent;
                    }
                }

                // Determine output path (ensure directory exists)
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LegendEntryTransparentForNegative.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

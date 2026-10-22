// Title: Reorder chart series in an Excel workbook by descending Y‑value totals using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads an Excel file with Aspose.Cells, computes the total of each series' Y‑values, sorts the series in descending order of those totals, removes the original series from the chart, and adds them back in the new sequence. | Write a C# routine that iterates through a chart's NSeries, extracts the Y‑range values, aggregates them, reorders the series based on the aggregated values, and saves the workbook with the updated series order.
// Common Searches: Aspose.Cells C# reorder chart series based on data totals | How to sort Excel chart series by sum of Y values using Aspose.Cells | Programmatically change series order in a chart with Aspose.Cells .NET | C# calculate total of series values and rearrange chart series in Excel | Aspose.Cells chart series sorting by descending Y‑value sum
// Tags: sort chart series Aspose.Cells | reorder Excel chart series .NET | calculate series Y values sum Aspose.Cells | modify NSeries collection C# | clear and add chart series Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, reads each series' name and cell ranges from the first chart, computes the sum of the Y‑values for every series, orders the series from highest to lowest sum, clears the existing series collection, re‑adds the series in the new order, and saves the workbook.
class ReorderChartSeries
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Collect series information: name, X range, Y range, and sum of Y values
            var seriesInfo = new List<(string Name, string XRange, string YRange, double Sum)>();

            foreach (Series series in chart.NSeries)
            {
                string name = series.Name;
                string xRange = series.XValues;
                string yRange = series.Values;

                double sum = 0.0;

                // Parse the Y range (assumes the range refers to the same worksheet)
                try
                {
                    // Use overload that includes the sheet name
                    CellArea area = CellArea.CreateCellArea(sheet.Name, yRange);
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            object val = sheet.Cells[row, col].Value;
                            if (val is double d)
                            {
                                sum += d;
                            }
                            else if (double.TryParse(val?.ToString(), out double parsed))
                            {
                                sum += parsed;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to parse range \"{yRange}\": {ex.Message}");
                }

                seriesInfo.Add((name, xRange, yRange, sum));
            }

            // Sort series so that the highest sum comes first
            seriesInfo.Sort((a, b) => b.Sum.CompareTo(a.Sum));

            // Remove all existing series from the chart (iterate backwards)
            for (int i = chart.NSeries.Count - 1; i >= 0; i--)
            {
                chart.NSeries.RemoveAt(i);
            }

            // Add series back in the new order
            foreach (var info in seriesInfo)
            {
                // Add series values; Add returns the index of the new series
                int index = chart.NSeries.Add(info.YRange, true);
                Series newSeries = chart.NSeries[index];
                newSeries.XValues = info.XRange;
                newSeries.Name = info.Name;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

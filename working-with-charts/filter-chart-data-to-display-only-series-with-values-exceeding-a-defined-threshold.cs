// Title: C# example: Remove Excel chart series whose maximum value is below a threshold using Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that iterates over a chart's NSeries, calculates each series' maximum cell value, and deletes the series if the max is less than a specified threshold. | Show how to create a column chart from a data range in Aspose.Cells and then filter out any series that do not meet a minimum value requirement. | Provide a C# snippet that loads an existing workbook, adds a chart, and programmatically removes series whose highest data point is under 50 using Aspose.Cells.
// Common Searches: Aspose.Cells C# filter chart series by maximum cell value | Remove low-value series from an Excel column chart using Aspose.Cells | Loop through chart series in Aspose.Cells and delete those below a threshold | C# keep only chart series with values greater than 50 in an Excel workbook
// Tags: filter chart series based on value Aspose.Cells C# | remove chart series by threshold Aspose.Cells | process NSeries Aspose.Cells | max value condition for chart series Aspose.Cells | column chart data range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// The program loads Input.xlsx, adds a column chart on the 'Data' worksheet, evaluates the maximum value of each series, removes any series whose maximum does not exceed 50, and saves the modified workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the source data
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet \"Data\" not found.");
                return;
            }

            // Add a column chart to the worksheet
            int chartIdx = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = dataSheet.Charts[chartIdx];

            // Define the data ranges
            chart.NSeries.Add("Data!$B$2:$D$10", true);               // series values
            chart.NSeries.CategoryData = "Data!$A$2:$A$10";          // categories

            // Threshold: only series whose maximum value exceeds this will be kept
            double threshold = 50.0;

            // Iterate through the series in reverse order so we can safely remove items
            for (int i = chart.NSeries.Count - 1; i >= 0; i--)
            {
                var series = chart.NSeries[i];

                // Get the cell range that holds the series values (strip sheet name if present)
                string seriesRange = series.Values; // e.g., "Data!$B$2:$B$10"
                string address = seriesRange.Contains("!") ? seriesRange.Split('!')[1] : seriesRange;

                AsposeRange range;
                try
                {
                    range = dataSheet.Cells.CreateRange(address);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create range \"{address}\": {ex.Message}");
                    continue;
                }

                // Determine the maximum value in the series
                double maxValue = double.MinValue;
                foreach (Cell cell in range)
                {
                    if (cell.Value is double d)
                    {
                        if (d > maxValue) maxValue = d;
                    }
                    else if (cell.Value is int iVal)
                    {
                        double d2 = iVal;
                        if (d2 > maxValue) maxValue = d2;
                    }
                }

                // Remove the series if its maximum does not exceed the threshold
                if (maxValue <= threshold)
                {
                    chart.NSeries.RemoveAt(i);
                }
            }

            // Save the workbook with the filtered chart
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Export Excel chart series values to a JSON file using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells example for extracting chart series data and saving as JSON | how to convert Excel chart NSeries to JSON with Aspose.Cells library | retrieve chart series range values from a workbook and serialize to JSON in C# | Aspose.Cells code to export chart data points to a JSON file | export first chart's series values from an .xlsx to JSON using C#
// Tags: Aspose.Cells export chart series to JSON | C# serialize Excel chart NSeries | retrieve chart values range Aspose.Cells | write JSON file from chart data C# | Excel chart data extraction Aspose.Cells | convert chart series values to JSON C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an Excel workbook with Aspose.Cells, accesses the first worksheet and its first chart, iterates through each series in the chart, resolves the series' value range, collects the cell values, builds an object containing the series name and its values, aggregates all series objects, serializes the collection to an indented JSON string, and writes the result to a specified JSON file while handling missing files, absent charts, and range errors.
class ChartDataExporter
{
    static void Main()
    {
        // Paths for input workbook and output JSON
        string workbookPath = "input.xlsx";
        string jsonPath = "chartData.json";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Assume the first chart on the worksheet is the target chart
            Chart chart = sheet.Charts[0];

            // Prepare a list to hold each series data
            var chartSeriesData = new List<object>();

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                try
                {
                    // Retrieve the range reference for the series values (e.g., "Sheet1!$B$2:$B$5")
                    string valuesRef = series.Values;

                    // Resolve the range reference to an actual Range object
                    AsposeRange valuesRange = null;
                    try
                    {
                        valuesRange = sheet.Cells.CreateRange(valuesRef);
                    }
                    catch (Exception exRange)
                    {
                        Console.WriteLine($"Warning: Unable to create range '{valuesRef}' - {exRange.Message}");
                        continue;
                    }

                    // Collect the cell values into a list
                    var values = new List<object>();
                    foreach (Cell cell in valuesRange)
                    {
                        values.Add(cell.Value);
                    }

                    // Build an anonymous object representing this series
                    var seriesObject = new
                    {
                        Name = series.Name,   // Series name (if set)
                        Values = values
                    };

                    chartSeriesData.Add(seriesObject);
                }
                catch (Exception exSeries)
                {
                    Console.WriteLine($"Warning: Failed to process a series - {exSeries.Message}");
                }
            }

            // Serialize the collected series data to JSON
            string json = JsonSerializer.Serialize(chartSeriesData, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write JSON to a file
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Chart data exported to {jsonPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

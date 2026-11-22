// Title: Export Excel chart series values to a formatted JSON file using Aspose.Cells in C#
// Common Searches: how to extract chart series values from an Excel file and save as JSON with Aspose.Cells C# | C# code to iterate Aspose.Cells chart NSeries and write data to a JSON array | exporting Excel chart data to JSON using Aspose.Cells library | convert Excel chart series to JSON file in .NET Core | Aspose.Cells example for reading chart source range and serializing to JSON
// Tags: Aspose.Cells export chart series to JSON | C# read chart NSeries values | Excel chart data serialization with Aspose.Cells | range creation from series.Values Aspose.Cells | JSON file output for Excel chart data | Aspose.Cells chart data extraction example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The sample loads an Excel workbook, accesses the first worksheet and its first chart, iterates through each series in the chart, creates a range from the series.Values reference, collects the cell values, builds a list of objects containing the series name and its values, serializes this list to an indented JSON string, and writes the result to a file named chartData.json.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chartData.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // List to hold each series' data
            var seriesList = new List<object>();

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                try
                {
                    // The series.Values property contains the source range reference (e.g., "Sheet1!$B$2:$B$6")
                    string valuesRef = series.Values;

                    // If the reference is empty, skip this series
                    if (string.IsNullOrEmpty(valuesRef))
                        continue;

                    // Create a Range object from the reference (fully qualified to avoid ambiguity)
                    Aspose.Cells.Range range = worksheet.Cells.CreateRange(valuesRef);

                    // Collect the cell values from the range
                    var values = new List<object>();
                    foreach (Cell cell in range)
                    {
                        values.Add(cell.Value);
                    }

                    // Build an anonymous object representing the series
                    var seriesObject = new
                    {
                        Name = series.Name,
                        Values = values
                    };

                    seriesList.Add(seriesObject);
                }
                catch (Exception exSeries)
                {
                    Console.WriteLine($"Warning: Failed to process a series. {exSeries.Message}");
                }
            }

            // Serialize the series list to a formatted JSON string
            string json = JsonSerializer.Serialize(seriesList, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Chart data exported to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

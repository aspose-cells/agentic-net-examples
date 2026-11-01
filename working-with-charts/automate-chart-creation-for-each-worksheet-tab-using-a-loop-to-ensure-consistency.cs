// Title: Create a column chart on every worksheet of an Excel file using Aspose.Cells for .NET in a loop
// Common Searches: C# loop through all worksheets and add a column chart with Aspose.Cells | Aspose.Cells generate a chart for each sheet programmatically | How to set chart title per worksheet using Aspose.Cells .NET
// Tags: add column chart to each worksheet Aspose.Cells | loop over workbook worksheets create charts C# | bind chart series to cell range Aspose.Cells | dynamic chart title per sheet Aspose.Cells | save Excel workbook with charts .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, iterates over every worksheet, adds a column chart positioned between rows 5‑20 and columns 0‑10, binds the series to the B2:B5 range (optionally sets category data from A2:A5), assigns a title that includes the worksheet name, and saves the updated file as a new workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate over each worksheet and add a column chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add a column chart (type, upper‑left row, upper‑left column, lower‑right row, lower‑right column)
                    int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIdx];

                    // Add the data series (values)
                    int seriesIdx = chart.NSeries.Add($"{sheet.Name}!$B$2:$B$5", true);

                    // Optionally set the category (X‑axis) data for the series
                    // Note: In some Aspose.Cells versions the CategoryData property may not be available.
                    // If needed, you can uncomment the line below and ensure your Aspose.Cells version supports it.
                    // chart.NSeries[seriesIdx].CategoryData = $"{sheet.Name}!$A$2:$A$5";

                    // Set a title for the chart
                    chart.Title.Text = $"Chart for {sheet.Name}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime errors and display a message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

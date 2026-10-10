// Title: Add a new data series to an existing Excel chart from a worksheet range using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an existing workbook, selects the first chart, adds a new series whose Y‑values come from a specified worksheet range, assigns a custom name to the series, and saves the workbook. | Update an Aspose.Cells chart to bind a new series to a dynamic cell range and change the series name without recreating the chart.
// Common Searches: asp.net add series to existing chart using Aspose.Cells C# | how to use NSeries.Add with a worksheet range in Aspose.Cells | set custom name for chart series programmatically Aspose.Cells | bind multiple data series to Excel chart from cell ranges Aspose.Cells C# | load workbook, modify chart data source and save with Aspose.Cells
// Tags: add chart series using NSeries.Add Aspose.Cells | bind chart series to worksheet range C# | set series name Aspose.Cells chart | update existing Excel chart data source Aspose.Cells | load and save workbook with modified chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel file, accesses the first worksheet and its first chart, adds a new series using the cell range B2:B5 for Y‑values, sets the series name to "New Series", and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Define the range that holds the Y‑values for the new series (e.g., B2:B5)
            string dataRange = $"{worksheet.Name}!$B$2:$B$5";

            // Add a new series to the chart using the worksheet range as the data source.
            // The Add method returns the index of the newly added series.
            int seriesIndex = chart.NSeries.Add(dataRange, false);
            Series newSeries = chart.NSeries[seriesIndex];
            newSeries.Name = "New Series";

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

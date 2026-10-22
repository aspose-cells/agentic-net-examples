// Title: Enable cell‑based data labels for a chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, retrieves the first chart series, and sets its DataLabels.ShowCellRange property to true. | Write a C# example that creates a column chart when none exists on a worksheet and activates cell‑based data labels for every series using Aspose.Cells. | Provide a C# snippet that iterates through all series in an existing Aspose.Cells chart and enables ShowCellRange for each series' data labels.
// Common Searches: Aspose.Cells C# set chart series data labels to use cell range | Enable cell based data labels for Excel chart series with Aspose.Cells .NET | How to turn on ShowCellRange for a series in an Aspose.Cells chart | C# Aspose.Cells add column chart and activate cell range labels programmatically
// Tags: Aspose.Cells ShowCellRange property | chart series data labels Aspose.Cells | create column chart programmatically Aspose.Cells | load workbook modify chart Aspose.Cells | C# Aspose.Cells update chart series

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing workbook, obtains or creates the first column chart on the first worksheet, ensures a series is present, sets Series.DataLabels.ShowCellRange to true, and saves the workbook with the updated chart labels.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first chart on the sheet, or create one if none exist
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                // Add a new column chart and retrieve its instance
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                chart = sheet.Charts[chartIndex];
            }

            // Ensure the chart has at least one series; add a dummy series if necessary
            if (chart.NSeries.Count == 0)
            {
                // Add a series that references a range (adjust as needed)
                chart.NSeries.Add("A1:A5", true);
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            // Activate cell‑based data labels for this series
            series.DataLabels.ShowCellRange = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

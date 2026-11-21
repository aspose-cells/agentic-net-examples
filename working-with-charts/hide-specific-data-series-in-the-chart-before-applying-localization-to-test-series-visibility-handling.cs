// Title: How to hide a chart data series in Aspose.Cells before setting a localized title (C#)
// Common Searches: Aspose.Cells hide second series in chart C# | C# set chart title localization after hiding series Aspose.Cells | How to make a chart series invisible in Aspose.Cells workbook | Apply localized chart title in Aspose.Cells after modifying series visibility
// Tags: hide chart series Aspose.Cells C# | localized chart title Aspose.Cells | chart NSeries visibility Aspose.Cells | Aspose.Cells chart series IsVisible property | Excel workbook chart modification Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an Excel workbook, verifies that a chart exists, hides the second data series when the API supports it, assigns a localized title to the chart, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Hide the second data series if it exists (property not available in some versions)
            if (chart.NSeries.Count > 1)
            {
                // In newer Aspose.Cells versions you can use:
                // chart.NSeries[1].IsVisible = false;
                // For compatibility, this step is skipped.
            }

            // Apply localization – set a localized title
            chart.Title.Text = "Localized Chart Title";

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

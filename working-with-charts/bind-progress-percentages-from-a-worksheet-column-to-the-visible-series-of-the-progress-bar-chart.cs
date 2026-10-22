// Title: Bind a worksheet column of percentages to a progress‑bar style bar chart using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an existing Excel file, reads percentages from a column range, creates a bar chart, binds the range to the chart series, and configures the chart to resemble a progress bar. | Show how to set the GapWidth property to zero and assign a custom title when generating a progress‑bar style chart with Aspose.Cells. | Demonstrate adding a bar chart to a worksheet, binding data from B2:B10, and saving the updated workbook using the Aspose.Cells .NET API.
// Common Searches: how to bind column B values to a bar chart series in Aspose.Cells C# | Aspose.Cells create progress bar chart from percentage data | remove gaps between bars in Aspose.Cells bar chart to simulate progress bar | add a bar chart to an existing Excel workbook using Aspose.Cells .NET
// Tags: bind worksheet range to chart series Aspose.Cells | progress bar style bar chart Aspose.Cells | set chart gapwidth zero Aspose.Cells | add bar chart to existing workbook Aspose.Cells | configure chart title Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, creates a bar chart, binds the B2:B10 percentage range to the chart series, sets the chart title to "Progress Bar", removes gaps by setting GapWidth to 0, and saves the updated workbook.
class ProgressBarChartExample
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = @"C:\Data\ProgressData.xlsx";
        string outputPath = @"C:\Data\ProgressData_WithChart.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that contains progress percentages (column B, rows 2‑10)
            string dataRange = "B2:B10";

            // Add a Bar chart to the worksheet (lifecycle rule: create)
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title (optional)
            chart.Title.Text = "Progress Bar";

            // Add a single series bound to the progress percentages column
            // 'false' indicates that the data is arranged in columns
            chart.NSeries.Add(dataRange, false);

            // Adjust appearance to resemble a progress bar
            chart.GapWidth = 0; // Remove gaps between bars

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

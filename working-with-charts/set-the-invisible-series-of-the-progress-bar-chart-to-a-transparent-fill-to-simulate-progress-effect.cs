// Title: Make the second series of an Excel progress bar chart transparent using Aspose.Cells for .NET
// Common Searches: how to set transparent fill for a specific series in an Aspose.Cells chart C# | Aspose.Cells C# hide remaining portion of progress bar chart by making series invisible | change series fill color to transparent in Excel chart with Aspose.Cells .NET | remove border and make series transparent in a progress bar chart using Aspose.Cells | set invisible series fill to transparent in a stacked column chart Aspose.Cells C#
// Tags: set series fill transparent Aspose.Cells .NET | hide chart series border C# | progress bar chart series styling Aspose.Cells | modify Excel chart series fill color .NET | transparent invisible series Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, locates the first chart, and sets the fill of the second series (used as the invisible part of a progress bar) to transparent while disabling its border, then saves the updated file.
class ProgressBarChartUpdater
{
    static void Main()
    {
        try
        {
            const string inputPath = "ProgressBar.xlsx";
            const string outputPath = "ProgressBar_Updated.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the progress bar chart
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Assume the chart is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Assume the chart we need to modify is the first chart
            Chart chart = sheet.Charts[0];

            // Iterate through all series in the chart
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];

                // Identify the invisible series (used for the "remaining" part of the bar)
                // Here we assume the second series (index 1) is the invisible one.
                if (i == 1)
                {
                    // Set the fill of the invisible series to transparent
                    series.Area.ForegroundColor = Color.Transparent;

                    // Hide the border to avoid any outline
                    series.Border.IsVisible = false;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

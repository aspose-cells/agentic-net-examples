// Title: Set data label position to InsideEnd for column series in an Excel chart with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set column chart data label position to InsideEnd | How to place data labels inside end of Excel column series using Aspose.Cells | Enable and position data labels for column charts in .NET with Aspose.Cells | Aspose.Cells example for showing data labels inside column chart series | Change data label placement for column chart series in C# Excel workbook
// Tags: Aspose.Cells set data label position | C# Excel column chart data labels | InsideEnd label placement Aspose.Cells | modify chart series labels .NET | Excel column chart label positioning | Aspose.Cells chart series label visibility

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, checks for a column or stacked column chart, enables data labels for each series, sets their position to InsideEnd, and saves the updated file.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Apply data label settings only to column-type charts
            if (chart.Type == ChartType.Column || chart.Type == ChartType.ColumnStacked)
            {
                foreach (Series series in chart.NSeries)
                {
                    // Enable data labels for the series
                    series.DataLabels.ShowValue = true;

                    // Optional: set the data label position if the enum is available
                    // series.DataLabels.Position = DataLabelPositionType.InsideEnd;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Change a specific chart series to a line type while keeping other series as columns using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, selects the first chart on the first worksheet, sets the second data series to ChartType.Line, ensures the first series stays ChartType.Column, and saves the modified file. | Generate a C# snippet using Aspose.Cells to convert only one series of a chart to a line chart while preserving the column chart type for the remaining series.
// Common Searches: Aspose.Cells change only second series to line chart in C# | How to set different chart types for individual series using Aspose.Cells .NET | C# Aspose.Cells keep column series and convert another series to line | Modify chart series type without affecting other series Aspose.Cells | Example code for mixed column and line chart series in Aspose.Cells
// Tags: Aspose.Cells set individual series chart type | C# change chart series to line | mixed column and line chart Aspose.Cells | chart series type manipulation .NET | Excel chart series conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads input.xlsx, accesses the first worksheet's first chart, forces the first series to a column chart and the second series to a line chart, then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Ensure the chart has at least two series
            if (chart.NSeries.Count < 2)
            {
                Console.WriteLine("The chart does not contain enough series to modify.");
                return;
            }

            // Change the second series to a line chart
            chart.NSeries[1].Type = ChartType.Line;

            // Explicitly set the first series to column type
            chart.NSeries[0].Type = ChartType.Column;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

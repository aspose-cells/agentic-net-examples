// Title: Update Excel chart title, change series fill color, and refresh the chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example to modify chart title and redraw chart | how to change series fill color in an Excel chart with Aspose.Cells and refresh it | C# code for calling Chart.Refresh after updating chart labels in Aspose.Cells | programmatically refresh an Excel chart after editing title and series color using .NET
// Tags: Aspose.Cells chart title modification C# | Chart.Refresh method Aspose.Cells .NET | change series fill color Aspose.Cells chart | update Excel chart labels Aspose.Cells | refresh Excel chart display C# Aspose

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example demonstrates how to load an Excel workbook with Aspose.Cells, access the first chart, update its title, change the fill color of the first data series, invoke Chart.Refresh to redraw the chart, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Modify the chart title
            chart.Title.Text = "Updated Chart Title";

            // Change fill color of the first series if it exists
            if (chart.NSeries.Count > 0)
            {
                chart.NSeries[0].Area.ForegroundColor = Color.LightBlue;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

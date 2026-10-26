// Title: Place the chart legend at the bottom and show its border with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to move chart legend to the bottom of an Excel chart | Aspose.Cells example to enable legend border visibility in a .NET workbook | code to set legend position and border color for an Excel chart using Aspose.Cells | programmatically change chart legend formatting with Aspose.Cells C# | Aspose.Cells chart legend bottom placement and border styling tutorial
// Tags: Aspose.Cells chart legend position bottom | Aspose.Cells enable legend border visibility | C# modify Excel chart legend formatting | Aspose.Cells LegendPositionType enum usage | Aspose.Cells set legend border color

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel file, accesses the first chart, moves its legend to the bottom, makes the legend border visible with a black color, and saves the workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
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

            // Set the legend to appear at the bottom of the chart
            // Use LegendPositionType enum (compatible with current Aspose.Cells versions)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Enable the border of the legend and set its color (optional)
            chart.Legend.Border.IsVisible = true;
            chart.Legend.Border.Color = Color.Black;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

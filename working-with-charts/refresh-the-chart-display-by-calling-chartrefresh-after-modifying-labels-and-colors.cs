// Title: Refresh an Aspose.Cells chart after updating series label and fill color using C#
// Common Searches: Aspose.Cells C# refresh chart after changing series name and fill color | call Chart.Refresh in Aspose.Cells to apply series label updates | C# update Excel chart series label and color then force redraw with Aspose.Cells | how to make chart reflect series color change without reopening workbook in Aspose.Cells | force chart redraw after modifying series properties in Aspose.Cells .NET
// Tags: Aspose.Cells Chart.Refresh usage | modify chart series label Aspose.Cells | set chart series fill color Aspose.Cells | C# refresh Excel chart Aspose.Cells | update first chart series Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet and its first chart, updates the first series name to "Updated Series" and changes its fill color to orange, calls Chart.Refresh to redraw the chart, and then saves the workbook.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Modify the first series if it exists
            if (chart.NSeries.Count > 0)
            {
                // Update series name (label)
                chart.NSeries[0].Name = "Updated Series";

                // Change the series fill color
                chart.NSeries[0].Area.ForegroundColor = Color.Orange;
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

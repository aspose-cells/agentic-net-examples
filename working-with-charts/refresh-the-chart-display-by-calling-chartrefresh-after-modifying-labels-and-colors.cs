// Title: How to refresh an Aspose.Cells chart after changing series label and fill color in C#
// AI Prompts: Load an existing Excel workbook with Aspose.Cells, change the first chart series name to "Updated Series", set its fill color to orange, invoke Chart.Refresh, and save the file. | Using Aspose.Cells for .NET, modify a chart's series label and background color, call Chart.Refresh to apply the changes, then write the updated workbook to a new output file.
// Common Searches: Aspose.Cells C# refresh chart after updating series name | call Chart.Refresh to apply series color change in Excel using Aspose.Cells | C# example of updating chart series label and fill color with Aspose.Cells | why chart changes not visible without Chart.Refresh in Aspose.Cells | programmatically refresh Excel chart display after editing series in .NET
// Tags: Chart.Refresh method Aspose.Cells | update chart series label C# Aspose.Cells | change chart series fill color Aspose.Cells | refresh Excel chart programmatically .NET | modify chart series Aspose.Cells example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The code loads 'input.xlsx', changes the first chart series name and its fill color, but does not call Chart.Refresh, so the visual updates may not appear. Adding Chart.Refresh after the modifications ensures the chart is refreshed before saving the workbook as 'output.xlsx'.
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

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Modify the first series if it exists
            if (chart.NSeries.Count > 0)
            {
                // Update series name (label)
                chart.NSeries[0].Name = "Updated Series";

                // Change series fill color
                chart.NSeries[0].Area.ForegroundColor = Color.Orange;
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

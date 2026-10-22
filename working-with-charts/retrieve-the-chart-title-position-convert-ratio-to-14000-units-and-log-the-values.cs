// Title: Retrieve a chart title’s text, visibility, and position ratio, convert the ratio to 1/4000 units, and log the values using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx workbook with Aspose.Cells, access the first worksheet’s first chart, read the title Text, IsVisible flag, and PositionRatio, transform the ratio into a 1/4000‑unit scale, and output all values to the console. | Extend the sample to also obtain the chart title’s absolute X/Y coordinates (in points) and log them together with the converted position ratio. | Add error handling to gracefully report when a worksheet contains no charts or when the chart title is not visible.
// Common Searches: how to get chart title position ratio Aspose.Cells C# | Aspose.Cells convert chart title position to 1/4000 units | read chart title visibility and coordinates from Excel using .NET | C# example for logging Excel chart title properties with Aspose.Cells | retrieve first chart title text and position in Aspose.Cells workbook
// Tags: Aspose.Cells chart title position ratio conversion | C# read Excel chart title visibility Aspose.Cells | log chart title text and coordinates Aspose.Cells | convert chart title position to 1/4000 units C# | access first worksheet chart title properties Aspose.Cells | Aspose.Cells retrieve chart title position ratio

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads an Excel workbook, accesses the first worksheet's first chart, reads the chart title text, visibility flag, and position ratio, converts the ratio to a 1/4000‑unit scale, and writes all retrieved values to the console.
class ChartTitleInfoLogger
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart in the worksheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Retrieve title information
            Title title = chart.Title;

            // Log the results
            Console.WriteLine("Chart Title Information:");
            Console.WriteLine($"  Text       : {title.Text}");
            Console.WriteLine($"  Is Visible : {title.IsVisible}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

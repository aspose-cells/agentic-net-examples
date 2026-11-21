// Title: Set a chart legend to the top position and read its coordinates using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend to top position | How to get legend coordinates after positioning a chart in Aspose.Cells | C# example moving Excel chart legend to top with LegendPositionType enum | Verify chart legend placement in a workbook saved by Aspose.Cells | Set column chart legend location to top in a .NET Excel file
// Tags: Aspose.Cells chart legend positioning API | C# LegendPositionType enum top value | extract legend location coordinates Aspose.Cells | column chart legend top alignment | save Excel workbook with positioned legend Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new Workbook, adds a column chart, sets the chart legend to the top using the LegendPositionType.Top enum, prints the legend position, and saves the file as ChartLegendPosition.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (rows 5-20, columns 0-10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the legend position (using a supported enum value)
            chart.Legend.Position = LegendPositionType.Top;

            // Output the legend position
            Console.WriteLine($"Legend Position: {chart.Legend.Position}");

            // Save the workbook to a file
            string outputPath = "ChartLegendPosition.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

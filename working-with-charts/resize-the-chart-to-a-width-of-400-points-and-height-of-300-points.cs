// Title: Resize an Aspose.Cells chart to 400 × 300 points using C#
// Common Searches: C# Aspose.Cells set chart width and height in points | How to change Excel chart size programmatically with Aspose.Cells | Aspose.Cells ChartObject dimensions example for .NET | Resize chart to 400x300 points using Aspose.Cells API | Set chartobject size in points in a workbook with Aspose.Cells
// Tags: Aspose.Cells chartobject resize points | set chart dimensions Aspose.Cells C# | adjust Excel chart size programmatically | chartobject width height property .NET | programmatic chart resizing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data and a column chart, then resizes the chart by setting its ChartObject Width to 400 points and Height to 300 points before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Resize the chart: width = 400 points, height = 300 points
            // ChartObject returns a ChartShape which exposes Width and Height properties (in points)
            chart.ChartObject.Width = 400;   // Set width to 400 points
            chart.ChartObject.Height = 300;  // Set height to 300 points

            // Optional: you can also set the position if needed
            // chart.ChartObject.X = 50;   // Horizontal offset from the left edge (points)
            // chart.ChartObject.Y = 50;   // Vertical offset from the top edge (points)

            // Save the workbook to a file
            string outputPath = "ResizedChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Chart resized to 400x300 points and workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}

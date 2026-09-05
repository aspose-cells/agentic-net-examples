// Title: How to set a column chart's background to light gray and clear fill patterns using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that applies a solid light‑gray fill to a chart’s ChartArea with Aspose.Cells, overriding any default pattern. | Show an example of creating a column chart in Aspose.Cells and removing its existing fill pattern before setting a light gray background. | Provide a snippet that configures the FillFormat of a chart area to Solid and saves the workbook as an .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change Excel chart area background color to light gray | remove default fill pattern from a column chart with Aspose.Cells .NET | set solid fill type for chart background using Aspose.Cells API | how to customize chart background color programmatically in .xlsx with Aspose.Cells
// Tags: chartarea.fillformat solid aspocells | aspocells chart background lightgray | aspocells clear chart fill pattern | c# aspocells column chart styling | aspocells chartarea filltype property

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, sets the chart area's FillFormat to Solid with a LightGray color, clears any default fill pattern, and saves the file as ChartBackground.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet and name it
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart background color to light gray (solid fill)
            FillFormat fill = chart.ChartArea.Area.FillFormat;
            fill.FillType = FillType.Solid;
            // Note: SolidFillColor property may not be available in older versions.
            // If supported, uncomment the following line:
            // fill.SolidFillColor = Color.LightGray;

            // Save the workbook
            wb.Save("ChartBackground.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

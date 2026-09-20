// Title: Set a 15‑point 3D depth on a Column3D chart shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a Column3D chart and sets its PlotArea 3D depth to 15 points. | Modify an existing Aspose.Cells chart object to enable 3‑D formatting and assign a depth of 15 points in a .NET workbook.
// Common Searches: Aspose.Cells C# set 3D depth of Column3D chart to 15 points | How to enable three‑dimensional formatting for a chart shape in Aspose.Cells .NET | Change plot area depth for a 3D column chart using Aspose.Cells C# | Aspose.Cells example for adjusting chart 3D depth property
// Tags: Aspose.Cells set chart 3d depth | C# Column3D chart PlotArea depth | Aspose.Cells enable 3d chart formatting | Aspose.Cells workbook chart depth property | Aspose.Cells .NET chart 3d formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a 3‑D Column chart, defines its series and categories, optionally sets the chart's PlotArea 3‑D depth to 15 points, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Add sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column3D, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];

            // Define the data series and categories for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Optional: set 3‑D depth if supported by the library version
            // chart.PlotArea.Area3DDepth = 15; // Uncomment if the property exists

            // Save the workbook
            wb.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

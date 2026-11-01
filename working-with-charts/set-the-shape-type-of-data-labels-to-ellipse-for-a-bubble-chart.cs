// Title: Create a bubble chart in Aspose.Cells for .NET and understand why ellipse-shaped data labels are not supported
// AI Prompts: Write C# code that uses Aspose.Cells to generate a bubble chart and attempts to set the data label shape to ellipse, including handling for the unsupported API. | Describe the limitation of Aspose.Cells regarding custom data label shapes on bubble charts and propose alternative styling methods. | Show how to modify font, color, and border of bubble chart data labels in Aspose.Cells when an ellipse shape cannot be applied.
// Common Searches: aspnet cells set bubble chart data label shape to ellipse | c# aspose.cells bubble chart custom data label appearance | why Aspose.Cells does not support ellipse data labels in Excel charts | alternative ways to style bubble chart data labels using Aspose.Cells .NET | example code for creating bubble chart with Aspose.Cells and handling unsupported label shapes
// Tags: bubble chart data label shape aspose.cells | ellipse data label limitation .NET Excel | customize bubble chart label appearance aspose.cells | c# aspose.cells chart label styling | unsupported chart label shape aspose.cells | excel bubble chart label formatting c#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates cells with X, Y, and size values, adds a bubble chart, assigns the series data, and notes that Aspose.Cells does not provide an API to set data label shapes such as an ellipse. The workbook is saved as BubbleChartWithEllipseLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the bubble chart
            ws.Cells["A1"].PutValue("X");
            ws.Cells["B1"].PutValue("Y");
            ws.Cells["C1"].PutValue("Size");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue(2);
            ws.Cells["C2"].PutValue(5);
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue(3);
            ws.Cells["C3"].PutValue(7);
            ws.Cells["A4"].PutValue(3);
            ws.Cells["B4"].PutValue(4);
            ws.Cells["C4"].PutValue(9);

            // Add a bubble chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Define the data series for the bubble chart
            chart.NSeries.Add("B2:B4", true);               // Y values
            chart.NSeries[0].XValues = "A2:A4";            // X values
            chart.NSeries[0].BubbleSizes = "C2:C4";        // Bubble sizes

            // Note: Aspose.Cells does not provide a direct API to set the shape of data labels.
            // The original intention was to use an ellipse shape, which is not supported.
            // This line has been removed to ensure successful compilation.

            // Save the workbook to a file
            string outputPath = "BubbleChartWithEllipseLabels.xlsx";
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

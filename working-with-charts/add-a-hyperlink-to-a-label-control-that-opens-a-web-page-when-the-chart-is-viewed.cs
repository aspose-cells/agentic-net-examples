// Title: Add a clickable hyperlink label to an Aspose.Cells chart in C#
// AI Prompts: Generate C# code that creates a rectangle shape on a worksheet, sets its text, and attaches a hyperlink so the label opens a web page when the chart is displayed using Aspose.Cells. | Show how to embed a hyperlink‑enabled label into an Excel chart created with Aspose.Cells for .NET. | Provide a step‑by‑step example that adds a shape with a URL to a chart worksheet and saves the workbook.
// Common Searches: how to attach a URL to a shape in an Aspose.Cells generated chart using C# | Aspose.Cells C# add hyperlink to chart label rectangle | example of adding clickable label to Excel chart with Aspose.Cells .NET | C# code to place a hyperlink shape on a worksheet that appears over a chart | Aspose.Cells add hyperlink to shape on chart worksheet
// Tags: Aspose.Cells add shape hyperlink C# | Excel chart label hyperlink Aspose.Cells | C# rectangle shape with URL Aspose.Cells | hyperlink on chart overlay shape Aspose.Cells | Aspose.Cells chart annotation hyperlink .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, fills cells with data, adds a column chart, inserts a rectangle shape labeled “Click Here”, assigns a hyperlink to the shape, and saves the file as ChartWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Add a rectangle shape that will act as a label with hyperlink
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape labelShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 5, 5, 30, 100);
            labelShape.Text = "Click Here";

            // Assign a hyperlink to the shape
            labelShape.AddHyperlink("https://www.example.com");

            // Save the workbook
            workbook.Save("ChartWithHyperlink.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

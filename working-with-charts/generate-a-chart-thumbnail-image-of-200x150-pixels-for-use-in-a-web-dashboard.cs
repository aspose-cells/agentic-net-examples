// Title: Generate a 200×150 px PNG thumbnail of a column chart using Aspose.Cells for .NET (C#)
// Description: C# code that builds a workbook, inserts sample data, creates a column chart, and saves the chart as a 200 × 150 pixel PNG thumbnail with ImageOrPrintOptions.SetDesiredSize (aspect ratio disabled), perfect for web dashboards or email reports.
// Keywords: Aspose.Cells | C# | .NET | chart thumbnail | PNG export | SetDesiredSize | ImageOrPrintOptions | column chart image | fixed pixel size | web dashboard chart
// Common Searches: Aspose.Cells export chart as 200x150 PNG | C# set exact pixel size for chart image Aspose.Cells | How to create chart thumbnail for web dashboard using Aspose.Cells | ImageOrPrintOptions SetDesiredSize example .NET | Generate fixed‑size chart image with Aspose.Cells
// Developer Intent: Produce a PNG chart thumbnail with exact 200 × 150 pixel dimensions for integration into a web dashboard or similar UI component.
// Use Cases: Render small, uniformly sized chart images for dashboard widgets. | Create preview thumbnails in a chart selector UI where layout consistency is required. | Export chart snapshots for email or PDF reports that need a predefined image size.
// AI Prompts: Show C# code to export an Aspose.Cells column chart as a 200x150 PNG thumbnail without preserving aspect ratio. | Explain how to configure ImageOrPrintOptions.SetDesiredSize for fixed‑size chart images in Aspose.Cells. | Provide a step‑by‑step example for generating a chart thumbnail suitable for a web dashboard using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// C# code that builds a workbook, inserts sample data, creates a column chart, and saves the chart as a 200 × 150 pixel PNG thumbnail with ImageOrPrintOptions.SetDesiredSize (aspect ratio disabled), perfect for web dashboards or email reports.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);
        chart.Title.Text = "Sample Chart";

        // Configure image options for a 200x150 thumbnail
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png
        };
        // Set exact pixel dimensions, do not preserve aspect ratio
        options.SetDesiredSize(200, 150, false);

        // Save the chart as an image file using the options
        chart.ToImage("chart_thumbnail.png", options);

        // (Optional) Save the workbook if needed
        // workbook.Save("chart_workbook.xlsx");
    }
}

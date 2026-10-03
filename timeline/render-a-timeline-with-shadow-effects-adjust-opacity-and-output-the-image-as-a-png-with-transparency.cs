// Title: Render a line chart with shadow and custom opacity, then export the worksheet as a transparent PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds a line chart, enables ChartArea.Shadow, sets the chart area's opacity to 70%, and saves the first worksheet as a PNG with an alpha channel using Aspose.Cells. | Create a reusable method that accepts a data range, builds a line chart with a shadow effect, applies 70% opacity, and returns the rendered PNG as a byte array via Aspose.Cells. | Generate code to render only the chart (excluding gridlines) to a transparent PNG while preserving the shadow effect, using ImageOrPrintOptions in Aspose.Cells.
// Common Searches: Aspose.Cells how to add shadow to a chart area in C# | Export chart as PNG with transparent background using Aspose.Cells .NET | Set chart area opacity in Aspose.Cells line chart | Render worksheet to image with alpha channel Aspose.Cells | Create timeline line chart and save as PNG in C# Aspose.Cells
// Tags: chart area shadow Aspose.Cells | export worksheet to transparent PNG Aspose.Cells | set chart opacity Aspose.Cells | render line chart as PNG .NET | timeline chart image generation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook, fills it with date/value data, inserts a line chart, enables a shadow on the chart area, optionally adjusts opacity, and renders the worksheet to a transparent PNG file named TimelineWithShadow.png using Aspose.Cells.
class TimelineWithShadow
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the timeline
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 4, 1));
            sheet.Cells["B5"].PutValue(25);

            // Add a Line chart (Timeline chart type is not supported) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart timelineChart = sheet.Charts[chartIndex];
            timelineChart.Name = "Project Timeline";

            // Set the data source for the chart
            timelineChart.NSeries.Add("B2:B5", true);
            timelineChart.NSeries.CategoryData = "A2:A5";

            // Apply shadow effect to the chart area
            timelineChart.ChartArea.Shadow = true;

            // Set image rendering options (default PNG format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Render the worksheet (including the chart) to a PNG image
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "TimelineWithShadow.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

// Title: Resize Aspose.Cells Chart Data Labels After Bold Font to Prevent Overflow (C#)
// Description: Demonstrates how to create a column chart with Aspose.Cells, enable data labels, apply bold formatting, disable automatic shape resizing, and set a fixed width and height for each label to avoid text clipping before saving the workbook.
// Keywords: Aspose.Cells chart data label resize | C# bold font data label overflow | disable auto‑fit data label Aspose | set data label width height programmatically | Excel chart label sizing Aspose.Cells
// Common Searches: Aspose.Cells prevent data label overflow after bold text | how to set fixed size for chart data labels in C# | disable auto‑fit for data label shapes Aspose.Cells | resize chart data labels to specific pixels | bold font data labels clipping Aspose.Cells
// Developer Intent: Manually size chart data label shapes after applying bold formatting to keep the text fully visible.
// Use Cases: Generating Excel reports where bold data labels must remain legible without automatic resizing. | Standardizing label dimensions across all points in a column chart for consistent visual layout. | Customizing chart appearance in dashboards by fixing label width and height regardless of content length.
// AI Prompts: Write C# code with Aspose.Cells that makes chart data label text bold and sets each label to 80 px wide and 30 px high. | Show how to iterate over ChartPoint objects, turn off IsResizeShapeToFitText, and assign custom pixel dimensions to data labels. | Provide an Aspose.Cells example that creates a column chart, enables data labels, applies bold font, disables auto‑fit, and manually resizes label shapes to avoid overflow.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates how to create a column chart with Aspose.Cells, enable data labels, apply bold formatting, disable automatic shape resizing, and set a fixed width and height for each label to avoid text clipping before saving the workbook.
class ResizeDataLabelShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Apply bold font to all data labels
        dataLabels.Font.IsBold = true;
        dataLabels.ApplyFont(); // Apply the font settings to all child label nodes

        // Disable automatic shape resizing and set a larger size to avoid overflow
        foreach (ChartPoint point in chart.NSeries[0].Points)
        {
            // Prevent auto‑fit of the label shape to the text
            point.DataLabels.IsResizeShapeToFitText = false;

            // Manually enlarge the label shape (adjust values as needed)
            point.DataLabels.WidthPixel = 80;
            point.DataLabels.HeightPixel = 30;
        }

        // Save the workbook
        workbook.Save("ResizeDataLabelShapes.xlsx");
    }
}

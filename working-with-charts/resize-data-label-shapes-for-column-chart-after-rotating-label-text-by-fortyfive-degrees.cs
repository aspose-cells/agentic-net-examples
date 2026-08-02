// Title: Resize Data Label Shapes After Rotating Text 45° in a Column Chart – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a column chart with Aspose.Cells, enable data labels, rotate the label text by 45 degrees, turn off automatic shape resizing, and set custom width (80 px) and height (30 px) for the label shapes before saving the workbook.
// Keywords: Aspose.Cells rotate data label | custom data label size .NET | disable data label auto resize | set data label width pixel | column chart label dimensions | C# Aspose.Cells chart labels | 45 degree label rotation
// Common Searches: Aspose.Cells set data label width and height after rotation | how to prevent data label auto‑resize in Aspose.Cells | rotate chart data labels 45 degrees C# | customize data label shape size Aspose.Cells column chart | Aspose.Cells chart label pixel dimensions
// Developer Intent: Manually size data label shapes after rotating the label text in a column chart.
// Use Cases: Generating Excel reports where angled data labels must stay within a fixed bounding box. | Designing dashboards with consistent label layout despite text rotation. | Creating printable charts where label clipping must be avoided by specifying exact shape dimensions.
// AI Prompts: Show C# code to rotate column chart data label text 45° and set label width to 80 pixels and height to 30 pixels using Aspose.Cells. | Explain how to disable automatic data label resizing and apply custom pixel dimensions after rotation in Aspose.Cells for .NET. | Provide step‑by‑step instructions for adjusting data label shape size to accommodate rotated text in a column chart.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a column chart with Aspose.Cells, enable data labels, rotate the label text by 45 degrees, turn off automatic shape resizing, and set custom width (80 px) and height (30 px) for the label shapes before saving the workbook.
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Set values
        chart.NSeries.CategoryData = "A2:A4";      // Set categories

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Rotate the data label text by 45 degrees
        dataLabels.RotationAngle = 45; // ChartTextFrame.RotationAngle

        // Disable automatic shape resizing so we can set a custom size
        dataLabels.IsResizeShapeToFitText = false;

        // Adjust the shape size to accommodate the rotated text (pixels)
        dataLabels.WidthPixel = 80;
        dataLabels.HeightPixel = 30;

        // Save the workbook
        workbook.Save("ResizedDataLabels.xlsx");
    }
}

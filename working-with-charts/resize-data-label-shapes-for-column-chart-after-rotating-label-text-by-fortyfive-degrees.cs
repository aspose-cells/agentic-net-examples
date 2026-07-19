// Title: Aspose.Cells .NET – Resize Column Chart Data Label Shapes After 45° Rotation
// Description: Demonstrates how to create a column chart, enable data labels, rotate the label text by 45 degrees, disable automatic shape resizing, and assign a fixed width (80 px) and height (30 px) to each label using Aspose.Cells for C#.
// Keywords: Aspose.Cells data label size | column chart label rotation C# | fixed pixel dimensions chart labels | disable auto resize data labels | set data label width height Aspose
// Common Searches: Aspose.Cells set data label width and height | rotate chart data labels 45 degrees .NET | prevent automatic resizing of Excel chart labels | custom size for column chart data labels | C# Aspose.Cells label shape dimensions
// Developer Intent: Apply a constant pixel size to rotated data label shapes in a column chart.
// Use Cases: Designing dashboards where rotated labels must stay within uniform boxes to avoid overlap. | Generating printable Excel reports that require consistent label dimensions regardless of text length. | Automating chart styling in bulk‑exported workbooks where label size control is essential.
// AI Prompts: How do I set a fixed pixel width and height for data labels after rotating them in Aspose.Cells for .NET? | Provide C# code that disables automatic resizing of chart data labels and assigns custom dimensions in a column chart. | Explain the steps to rotate column chart data labels 45° and keep each label shape at 80 × 30 pixels using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates how to create a column chart, enable data labels, rotate the label text by 45 degrees, disable automatic shape resizing, and assign a fixed width (80 px) and height (30 px) to each label using Aspose.Cells for C#.
class ResizeDataLabelShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the column chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Rotate the data label text by 45 degrees
        dataLabels.RotationAngle = 45;

        // Disable automatic shape resizing and set custom dimensions
        dataLabels.IsResizeShapeToFitText = false;
        dataLabels.WidthPixel = 80;   // Width of the label shape in pixels
        dataLabels.HeightPixel = 30;  // Height of the label shape in pixels

        // Save the workbook
        workbook.Save("ResizedDataLabels.xlsx");
    }
}

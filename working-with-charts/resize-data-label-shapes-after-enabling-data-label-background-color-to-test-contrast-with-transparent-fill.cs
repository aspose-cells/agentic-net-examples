// Title: Resize Chart Data Label Shapes & Apply Transparent Background with Aspose.Cells (C#)
// Description: Creates a workbook, adds a column chart, enables data labels, sets the label background to Transparent, disables automatic shape resizing, and defines a fixed 80 × 30 pixel rectangular label before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | resize data label shape | transparent label background | custom label size | disable auto resize | DataLabelShapeType.Rect | WidthPixel | HeightPixel | LabelPositionType.Center | column chart example | GitHub Aspose.Cells demo
// Common Searches: Aspose.Cells set data label width and height | how to make chart data label background transparent in C# | disable automatic resizing of data labels Aspose.Cells | custom shape for chart data labels Aspose.Cells | resize data label shapes pixel units Aspose.Cells
// Developer Intent: Define fixed dimensions and a transparent background for chart data label shapes in Aspose.Cells.
// Use Cases: Generate charts with uniformly sized labels regardless of text length. | Test visual contrast of data labels against varying series colors by using a transparent fill. | Create printable reports where label dimensions must remain constant for layout stability.
// AI Prompts: Show C# code using Aspose.Cells to set a fixed pixel width and height for chart data label shapes and make the background transparent. | Explain how to turn off automatic data label resizing and assign a rectangular shape type in an Aspose.Cells column chart. | Provide steps to evaluate label contrast by applying a transparent background while customizing label dimensions in Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a column chart, enables data labels, sets the label background to Transparent, disables automatic shape resizing, and defines a fixed 80 × 30 pixel rectangular label before saving the file.
class ResizeDataLabelShapesDemo
{
    static void Main()
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
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;
        dataLabels.Position = LabelPositionType.Center;

        // Set background mode to Transparent to test contrast with fill
        dataLabels.BackgroundMode = BackgroundMode.Transparent;

        // Disable automatic resizing so we can define a custom shape size
        dataLabels.IsResizeShapeToFitText = false;

        // Define custom size for the data label shape (pixel units)
        dataLabels.WidthPixel = 80;   // Width of the label
        dataLabels.HeightPixel = 30;  // Height of the label

        // Optionally set a visible shape type
        dataLabels.ShapeType = DataLabelShapeType.Rect;

        // Save the workbook to a file
        workbook.Save("ResizeDataLabelShapes.xlsx");
    }
}

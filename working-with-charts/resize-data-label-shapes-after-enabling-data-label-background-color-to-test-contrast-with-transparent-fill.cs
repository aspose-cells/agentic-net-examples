// Title: Resize Chart Data Label Shapes & Apply Transparent Background with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, insert a column chart, enable data labels, set a rectangular shape, make the background transparent, turn off automatic resizing, and define custom width and height in pixels using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart data labels | resize data label shape | transparent background | manual label size | DataLabelShapeType.Rect | BackgroundMode.Transparent | IsResizeShapeToFitText false | WidthPixel HeightPixel | column chart Excel automation
// Common Searches: how to set custom width for chart data labels Aspose.Cells | transparent background for Excel chart data labels .NET | disable auto resize of data label shapes in Aspose.Cells | change data label shape to rectangle in C# chart | Aspose.Cells example for manual data label sizing
// Developer Intent: The developer wants precise control over chart data label dimensions and appearance, disabling auto‑resize and using a transparent fill to evaluate label contrast.
// Use Cases: Create column charts with fixed‑size rectangular labels for a uniform layout across varying values. | Generate Excel reports where label backgrounds are transparent, allowing underlying series colors to show through. | Produce charts where auto‑sizing is turned off so custom fonts and label dimensions remain consistent.
// AI Prompts: Provide C# code to set a fixed pixel width and height for Aspose.Cells chart data labels and turn off automatic resizing. | Show how to apply a transparent background to rectangular data labels in an Aspose.Cells column chart. | Explain the steps to change font color after manually resizing data label shapes using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDataLabelResizeDemo
{
    // Shows how to build a workbook, insert a column chart, enable data labels, set a rectangular shape, make the background transparent, turn off automatic resizing, and define custom width and height in pixels using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series' data labels
            DataLabels dataLabels = chart.NSeries[0].DataLabels;

            // Show the value in each data label
            dataLabels.ShowValue = true;

            // Set a rectangular shape for the data label
            dataLabels.ShapeType = DataLabelShapeType.Rect;

            // Enable a transparent background to test contrast
            dataLabels.BackgroundMode = BackgroundMode.Transparent;

            // Disable automatic resizing of the shape to fit the text
            dataLabels.IsResizeShapeToFitText = false;

            // Manually set the size of the data label shape (in pixels)
            dataLabels.WidthPixel = 80;   // narrower than default
            dataLabels.HeightPixel = 30;  // shorter than default

            // Optionally change the font color to make the contrast visible
            dataLabels.Font.Color = Color.Black;
            dataLabels.Font.Size = 10;

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelResizeDemo.xlsx");
        }
    }
}

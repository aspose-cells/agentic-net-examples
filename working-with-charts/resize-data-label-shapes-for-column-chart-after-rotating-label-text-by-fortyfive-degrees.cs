// Title: C# – Resize Column Chart Data Label Shapes After 45° Rotation with Aspose.Cells
// Description: Demonstrates how to create a column chart in Aspose.Cells, enable data labels, rotate the label text 45 degrees, turn off automatic shape resizing, and assign fixed pixel width and height to each label before saving the workbook.
// Keywords: Aspose.Cells C# column chart | rotate data label text 45 degrees | custom data label size Aspose.Cells | disable auto resize data labels | WidthPixel HeightPixel chart | Excel chart label formatting .NET | set data label shape dimensions | Aspose.Cells chart customization | C# Excel automation data labels | fixed label width height Aspose
// Common Searches: Aspose.Cells set data label width and height after rotation | How to disable automatic resizing of chart data labels in C# | Rotate column chart data labels 45 degrees Aspose.Cells | Custom size for data label shapes in Aspose.Cells chart | C# code to fix data label dimensions in Excel chart
// Developer Intent: Apply a fixed pixel width and height to column‑chart data label shapes after rotating the label text 45 degrees using Aspose.Cells for .NET.
// Use Cases: Create a column chart with rotated data labels that remain legible on dashboards. | Prevent label clipping by disabling auto‑resize and defining explicit label dimensions. | Maintain a consistent visual style across multiple series when labels are angled.
// AI Prompts: Generate C# code with Aspose.Cells that rotates column chart data label text 45° and sets a fixed WidthPixel and HeightPixel for each label. | Explain step‑by‑step how to turn off automatic data label resizing and manually size label shapes after rotation in Aspose.Cells for .NET. | Provide a tutorial for creating a column chart with custom‑sized, rotated data labels using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart in Aspose.Cells, enable data labels, rotate the label text 45 degrees, turn off automatic shape resizing, and assign fixed pixel width and height to each label before saving the workbook.
    class ResizeDataLabelShapesDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook(FileFormatType.Xlsx);
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Cat 1");
            sheet.Cells["A3"].PutValue("Cat 2");
            sheet.Cells["A4"].PutValue("Cat 3");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;

            // Rotate the data label text by 45 degrees
            dataLabels.RotationAngle = 45;

            // Disable automatic resizing of the label shape to fit the rotated text
            dataLabels.IsResizeShapeToFitText = false;

            // Set a custom size for the data label shape (width and height in pixels)
            dataLabels.WidthPixel = 80;   // Adjust as needed
            dataLabels.HeightPixel = 30;  // Adjust as needed

            // Save the workbook
            workbook.Save("ResizeDataLabelShapesDemo.xlsx");
        }
    }
}

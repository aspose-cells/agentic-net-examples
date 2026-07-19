// Title: Aspose.Cells for .NET – Set Chart Data Label Shape (Triangle fallback)
// Description: Creates a workbook, adds sample data, inserts a column chart, enables data labels, and attempts to set the label shape to a triangle. Because Aspose.Cells' DataLabelShapeType enum lacks a triangle option, the code uses the Rect shape as the nearest alternative before saving the file.
// Keywords: Aspose.Cells | C# chart data label shape | DataLabelShapeType | custom chart label triangle | Excel chart label shape | Aspose.Cells .NET example | set data label shape
// Common Searches: Aspose.Cells set data label shape triangle | C# change chart label shape Aspose.Cells | DataLabelShapeType enum options | how to customize chart data labels in Aspose.Cells | fallback shape for missing triangle label
// Developer Intent: Apply a specific shape to chart data labels—preferably a triangle—and use the closest supported shape when the triangle option is unavailable, using Aspose.Cells in C#.
// Use Cases: Programmatically generate Excel charts with customized data label shapes for clearer visual emphasis. | Handle missing enum values by selecting an appropriate fallback shape (e.g., Rect) in automated report generation. | Demonstrate chart styling techniques in Aspose.Cells tutorials or documentation.
// AI Prompts: Generate C# code with Aspose.Cells that sets chart data label shape to triangle, and fallback to Rect if triangle is not supported. | Show how to enable data labels on a series and change their shape type in a custom Aspose.Cells chart. | Explain strategies for handling unavailable DataLabelShapeType values when customizing chart labels in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, enables data labels, and attempts to set the label shape to a triangle. Because Aspose.Cells' DataLabelShapeType enum lacks a triangle option, the code uses the Rect shape as the nearest alternative before saving the file.
class SetDataLabelShapeTriangle
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

        // Add a custom chart (using Column as an example)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series and enable data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Set the shape type of data labels.
        // The DataLabelShapeType enumeration does not contain a triangle shape,
        // so we use Rect as the closest available example.
        series.DataLabels.ShapeType = DataLabelShapeType.Rect;

        // Save the workbook
        workbook.Save("DataLabelShapeTriangle.xlsx");
    }
}

// Title: Aspose.Cells .NET: Set Line Chart Data Labels to Rounded Rectangle (C#)
// Description: Creates a workbook, adds sample data, inserts a line chart, enables data labels for the first series, changes the label shape to a rounded rectangle using DataLabelShapeType.RoundRect, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# line chart data labels | DataLabels.ShapeType RoundRect | rounded rectangle data label Aspose.Cells | set chart data label shape .NET | Aspose.Cells chart formatting | Excel line chart custom data labels
// Common Searches: Aspose.Cells set data label shape to rounded rectangle | C# line chart data label RoundRect Aspose.Cells | How to change DataLabels.ShapeType in Aspose.Cells | Rounded rectangle data labels for Excel chart .NET | Aspose.Cells chart label shape customization
// Developer Intent: Apply a rounded‑rectangle shape to data labels of a line‑chart series in Aspose.Cells for .NET.
// Use Cases: Enhance readability of line‑chart values by using rounded‑rectangle labels in automated reports. | Maintain a consistent label style across multiple charts for corporate branding. | Generate Excel files with custom‑shaped data labels to match design guidelines.
// AI Prompts: Show C# code to set DataLabels.ShapeType to RoundRect for a line chart using Aspose.Cells. | How can I apply different data label shapes to each series in an Aspose.Cells chart? | Provide an example of customizing data label appearance (font, color, shape) in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a line chart, enables data labels for the first series, changes the label shape to a rounded rectangle using DataLabelShapeType.RoundRect, and saves the file as an Excel workbook.
class SetDataLabelShapeType
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Set the shape type of data labels to rounded rectangle
        series.DataLabels.ShapeType = DataLabelShapeType.RoundRect;

        // Save the workbook
        workbook.Save("LineChartWithRoundedRectDataLabels.xlsx");
    }
}

// Title: Aspose.Cells C# – Set line chart data label shape to rounded rectangle
// Description: Creates a workbook, adds a line chart with sample data, enables data labels for the first series, changes the label shape to a rounded rectangle using DataLabelShapeType.RoundRect, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart example | line chart data labels | rounded rectangle label shape | DataLabelShapeType | Excel chart formatting | chart label customization
// Common Searches: Aspose.Cells set data label shape rounded rectangle | C# line chart data labels shape type | DataLabelShapeType.RoundRect example | change chart label shape Aspose.Cells | how to format data labels in Aspose.Cells chart
// Developer Intent: Apply a rounded‑rectangle shape to the data labels of a line‑chart series using Aspose.Cells for .NET.
// Use Cases: Enhance visual emphasis of line‑chart values by using rounded‑rectangle data labels. | Match corporate branding guidelines that require specific label shapes in generated Excel reports. | Create reusable chart templates where data label shapes are preset for consistency across multiple workbooks.
// AI Prompts: Generate C# code that sets DataLabels.ShapeType to RoundRect for a line chart series with Aspose.Cells. | Show an Aspose.Cells example that creates a line chart, enables data labels, and applies a rounded rectangle shape. | List all DataLabelShapeType options in Aspose.Cells and demonstrate how to switch between them in code.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a line chart with sample data, enables data labels for the first series, changes the label shape to a rounded rectangle using DataLabelShapeType.RoundRect, and saves the file as an Excel workbook.
class SetDataLabelShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data for the line chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Y values
        chart.NSeries.CategoryData = "A2:A4";      // X categories

        // Access the first series and enable data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Set the shape type of the data labels to rounded rectangle
        series.DataLabels.ShapeType = DataLabelShapeType.RoundRect;

        // Save the workbook to a file
        workbook.Save("LineChartDataLabelRoundedRect.xlsx");
    }
}

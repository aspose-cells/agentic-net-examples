// Title: Aspose.Cells for .NET – Set Chart Category Axis Tick Labels to Stacked (Vertical) Direction
// Description: Shows how to create a workbook, add sample data, insert a column chart, and apply ChartTextDirectionType.Stacked to the CategoryAxis.TickLabels so the tick labels appear vertically stacked before saving the Excel file.
// Keywords: Aspose.Cells | ChartTextDirectionType | Stacked | tick label direction | category axis | vertical labels | C# chart example | Excel chart label orientation | .NET Aspose.Cells | chart axis formatting
// Common Searches: Aspose.Cells set tick label direction | ChartTextDirectionType.Stacked C# | vertical stacked tick labels Aspose.Cells | change category axis label orientation .NET | Excel chart label direction Aspose.Cells example
// Developer Intent: Apply ChartTextDirectionType.Stacked to a chart’s CategoryAxis.TickLabels to display tick labels in a vertical stacked layout.
// Use Cases: Generate column charts where long category names need a compact, stacked vertical presentation for better readability. | Create automated Excel reports that enforce a consistent stacked label style across multiple charts in a workbook. | Design dashboards with narrow columns, using stacked tick labels to prevent label truncation and overlap.
// AI Prompts: Provide a C# snippet that sets both category and value axis tick labels to stacked orientation using Aspose.Cells. | Explain how ChartTextDirectionType.Stacked differs from Rotated and Horizontal in Aspose.Cells charts. | Show how to programmatically apply stacked tick label direction to all charts in an existing workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add sample data, insert a column chart, and apply ChartTextDirectionType.Stacked to the CategoryAxis.TickLabels so the tick labels appear vertically stacked before saving the Excel file.
class SetTickLabelsDirection
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set tick labels direction to Stacked (vertical stacked text) on the category axis
        chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Stacked;

        // Save the workbook
        workbook.Save("TickLabelsStackedDirection.xlsx");
    }
}

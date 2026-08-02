// Title: Apply Built‑In Chart Style and Freeze Data Rows with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills A1:B4 with categories and values, adds a column chart (rows 5‑20, cols 0‑8), sets the series and category data, applies built‑in style 2, freezes the first four rows via FreezePanes, and saves the file as ChartStyleAndFreezeDemo.xlsx.
// Keywords: Aspose.Cells C# chart style | Aspose.Cells freeze panes | apply built‑in chart style .NET | freeze rows worksheet Aspose | column chart style Aspose.Cells | Excel chart formatting C# | freeze top rows Aspose.Cells | chart appearance consistency
// Common Searches: Aspose.Cells set chart style programmatically | How to freeze rows in Aspose.Cells C# | Freeze panes while keeping chart data visible Aspose.Cells | Apply built‑in chart style to column chart Aspose.Cells | C# code to freeze header rows in Excel using Aspose.Cells
// Developer Intent: Programmatically apply a predefined chart style and lock the source data rows in place with FreezePanes.
// Use Cases: Generate reports with uniform chart formatting and fixed data headers. | Build interactive Excel dashboards where scrolling does not hide source rows. | Prepare workbooks for printing or sharing where chart style is consistent and top rows stay visible. | Automate template creation that requires specific chart appearance and frozen rows.
// AI Prompts: Write C# using Aspose.Cells to add a bar chart, apply built‑in style 4, and freeze the first five rows. | Generate a method that creates a line chart, sets style 7, and freezes rows 1‑10 with FreezePanes in Aspose.Cells for .NET. | Provide sample code to apply a chart style and freeze header rows for a pie chart in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills A1:B4 with categories and values, adds a column chart (rows 5‑20, cols 0‑8), sets the series and category data, applies built‑in style 2, freezes the first four rows via FreezePanes, and saves the file as ChartStyleAndFreezeDemo.xlsx.
class ChartStyleAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart covering rows 5‑20 and columns 0‑8
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data series and category data for the chart
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a built‑in chart style (style number must be between 1 and 48)
        chart.Style = 2;

        // Freeze the rows that contain the data (rows 1‑4) so they stay visible while scrolling
        // Freeze at row index 5 (first row after the data) with 5 frozen rows and 0 frozen columns
        sheet.FreezePanes(5, 0, 5, 0);

        // Save the workbook
        workbook.Save("ChartStyleAndFreezeDemo.xlsx");
    }
}

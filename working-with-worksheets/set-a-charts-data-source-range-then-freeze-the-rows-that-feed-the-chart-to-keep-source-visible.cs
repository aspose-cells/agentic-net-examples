// Title: Aspose.Cells for .NET – Set Chart Data Range and Freeze Source Rows/Columns
// Description: C# example that creates a workbook, fills A1:B10, adds a column chart, assigns the chart's data source with SetChartDataRange, then freezes the first 10 rows and first two columns using FreezePanes so the source stays visible while scrolling.
// Keywords: Aspose.Cells | .NET | C# chart data range | SetChartDataRange | FreezePanes | freeze rows Excel | freeze columns Excel | chart source visibility | column chart example | Excel worksheet freezing
// Common Searches: Aspose.Cells set chart data range C# | How to freeze panes for chart source in Aspose.Cells | Freeze rows and columns in Excel using Aspose.Cells .NET | Keep chart data visible while scrolling Aspose.Cells | SetChartDataRange and FreezePanes example
// Developer Intent: Assign a chart’s data source range and lock the source rows/columns in view with FreezePanes.
// Use Cases: Generate a report where chart data tables remain on‑screen as users scroll through large worksheets. | Create dashboards with multiple charts and keep each chart’s source range anchored for quick reference. | Prepare presentation‑ready Excel files where the underlying data is frozen to avoid accidental scrolling off‑screen.
// AI Prompts: Write C# code with Aspose.Cells to set a line chart’s data range to A1:C15 and freeze the first 15 rows and first three columns. | Show how FreezePanes parameters map to row and column indices when keeping chart source data visible in an Aspose.Cells workbook. | Provide a step‑by‑step explanation of using SetChartDataRange and FreezePanes together for chart visibility in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, fills A1:B10, adds a column chart, assigns the chart's data source with SetChartDataRange, then freezes the first 10 rows and first two columns using FreezePanes so the source stays visible while scrolling.
class ChartDataRangeAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (rows 1‑10, columns A‑B)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart covering rows 12‑22, columns C‑H (arbitrary position)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 2, 22, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source range (A1:B10) and plot by column (vertical = true)
        chart.SetChartDataRange("A1:B10", true);

        // Freeze the rows that contain the chart source data (first 10 rows) 
        // and the first two columns so the source stays visible while scrolling
        sheet.FreezePanes(10, 2, 10, 2);

        // Save the workbook
        workbook.Save("ChartDataRangeAndFreeze.xlsx");
    }
}

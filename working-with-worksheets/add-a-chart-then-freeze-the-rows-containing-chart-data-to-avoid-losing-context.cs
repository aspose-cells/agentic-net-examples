// Title: Add a Column Chart and Freeze Data Rows using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, fills rows 1‑11 with category/value data, inserts a column chart linked to that range, freezes the first 11 rows to keep the source table visible, and saves the file as ChartWithFrozenRows.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Add column chart | FreezePanes | Freeze rows Excel | Chart data source | Excel automation | Aspose.Cells example | GitHub sample
// Common Searches: Aspose.Cells add chart and freeze rows | C# freeze top rows after inserting chart | How to use FreezePanes with charts in Aspose.Cells | Create column chart in Excel with Aspose.Cells .NET | Aspose.Cells sample code for chart and FreezePanes
// Developer Intent: Insert a column chart into a worksheet and lock the rows that contain its source data.
// Use Cases: Sales dashboards where the data table stays visible while scrolling through a large chart. | Financial reports that need header rows frozen above a chart placed lower in the sheet. | Excel templates that automatically protect context by freezing rows after adding a chart.
// AI Prompts: Generate C# code with Aspose.Cells to add a line chart and freeze the first 5 rows. | Show an Aspose.Cells example that creates a pie chart and uses FreezePanes to keep legend rows visible. | Explain how to configure FreezePanes to lock both rows and columns around a chart in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, fills rows 1‑11 with category/value data, inserts a column chart linked to that range, freezes the first 11 rows to keep the source table visible, and saves the file as ChartWithFrozenRows.xlsx.
class AddChartAndFreezeRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart (rows 1‑11)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 11; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart that occupies rows 12‑30 and columns C‑I
        int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 2, 30, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source to the populated range
        chart.NSeries.Add("=Sheet1!$A$1:$B$11", true);

        // Freeze the rows that contain the chart data (rows 1‑11)
        // Row index 11 (zero‑based) is the first row that will NOT be frozen.
        // Freeze 11 rows and 0 columns.
        sheet.FreezePanes(11, 0, 11, 0);

        // Save the workbook
        workbook.Save("ChartWithFrozenRows.xlsx", SaveFormat.Xlsx);
    }
}

// Title: Add a Column Chart and Freeze Data Rows with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, populates rows 1‑11 with sample categories and values, inserts a column chart below the table, sets the chart's source to A1:B11, freezes the first 11 rows to keep the data visible while scrolling, and saves the file as ChartWithFrozenRows.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart C# | freeze panes Aspose.Cells | add column chart .NET | Aspose.Cells FreezePanes example | Excel export with chart C#
// Common Searches: Aspose.Cells add column chart and freeze rows | How to use FreezePanes after creating a chart in C# | Create chart and lock top rows Aspose.Cells | C# Aspose.Cells freeze rows example
// Developer Intent: Generate a worksheet that contains a column chart and keeps the source rows fixed with FreezePanes.
// Use Cases: Sales report where the data table stays in view while the user scrolls through the chart. | Financial dashboard that combines a chart with a locked data section for easy reference. | Automated Excel export that includes a visual chart and prevents the underlying data from scrolling out of sight.
// AI Prompts: Provide C# code to add a line chart and freeze the first 15 rows using Aspose.Cells. | Show how to add multiple charts and apply FreezePanes to each corresponding data range in Aspose.Cells. | Explain the zero‑based indexing of FreezePanes in Aspose.Cells and how to adjust parameters for different chart positions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a new workbook, populates rows 1‑11 with sample categories and values, inserts a column chart below the table, sets the chart's source to A1:B11, freezes the first 11 rows to keep the data visible while scrolling, and saves the file as ChartWithFrozenRows.xlsx using Aspose.Cells for .NET.
class AddChartAndFreezeRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart (rows 1‑11, zero‑based indexes 0‑10)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 10; i++)
        {
            sheet.Cells[i, 0].PutValue("Item " + i);   // Column A
            sheet.Cells[i, 1].PutValue(i * 5);        // Column B
        }

        // Add a column chart positioned below the data range
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIdx = sheet.Charts.Add(ChartType.Column, 12, 2, 30, 10);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data source for the chart (including header row)
        chart.NSeries.Add("=Sheet1!$A$1:$B$11", true);

        // Freeze the rows that contain the chart data (rows 1‑11)
        // FreezePanes(row, column, freezedRows, freezedColumns)
        // Row and column indexes are zero‑based; row = 11 splits after the 11th row.
        sheet.FreezePanes(11, 0, 11, 0);

        // Save the workbook
        workbook.Save("ChartWithFrozenRows.xlsx", SaveFormat.Xlsx);
    }
}

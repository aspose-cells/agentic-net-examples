// Title: Apply Built‑In Chart Style and Freeze Header Rows with Aspose.Cells in C#
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, apply a built‑in chart style (e.g., style #2), freeze the first four rows that contain the chart source data using Worksheet.FreezePanes, and save the file as ChartStyleAndFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart style C# | freeze panes Aspose.Cells | Worksheet.FreezePanes .NET | apply built‑in chart style | column chart Aspose.Cells | Excel automation C# | freeze header rows Excel | Aspose.Cells example
// Common Searches: Aspose.Cells set chart style and freeze rows | How to freeze panes after adding a chart in C# | Apply built‑in chart style Aspose.Cells .NET | Worksheet.FreezePanes example with chart data | Freeze header rows in Excel using Aspose.Cells
// Developer Intent: The developer wants to style a chart with a predefined Aspose.Cells chart style and lock the rows that provide the chart’s data so the layout remains consistent when scrolling.
// Use Cases: Create a financial report where the header rows stay visible while a styled column chart displays quarterly results. | Generate an automated Excel export that applies a corporate chart theme and freezes the data‑source rows to preserve dashboard layout. | Build a multi‑chart dashboard workbook that uses consistent styling and freezes source rows to prevent accidental scrolling off the data range.
// AI Prompts: Show C# code to apply a built‑in chart style and freeze specific rows with Aspose.Cells. | How do I use Worksheet.FreezePanes to lock the first four rows after creating a chart in Aspose.Cells for .NET? | Explain the parameters of Worksheet.FreezePanes and give an example that freezes header rows for a chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartStyleAndFreeze
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, apply a built‑in chart style (e.g., style #2), freeze the first four rows that contain the chart source data using Worksheet.FreezePanes, and save the file as ChartStyleAndFreeze.xlsx using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data that will be used by the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart that covers rows 5‑20 and columns 0‑8
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data series and category data for the chart
            chart.NSeries.Add("B2:B4", false);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a built‑in style (e.g., style #2) to the chart
            chart.Style = 2; // Valid values are 1‑48; -1 means not set

            // Freeze the rows that contain the chart data (rows 1‑4)
            // FreezePanes(rowIndex, columnIndex, freezedRows, freezedColumns)
            // Row and column indices are zero‑based, so row 4 corresponds to the 5th row.
            worksheet.FreezePanes(4, 0, 4, 0);

            // Save the workbook
            workbook.Save("ChartStyleAndFreeze.xlsx");
        }
    }
}

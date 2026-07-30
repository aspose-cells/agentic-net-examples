// Title: C# – Add Data Labels to a Column Chart and Freeze Top Rows with Aspose.Cells
// Description: Creates a workbook, fills A1:B4, inserts a column chart, enables data labels (values shown outside columns), freezes rows 1‑4 via FreezePanes, and saves as ChartWithDataLabelsAndFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | column chart | data labels | FreezePanes | freeze rows | chart label position | Excel export | chart series values
// Common Searches: Aspose.Cells show data labels on chart C# | Freeze top rows in Excel with Aspose.Cells | Set label position OutsideEnd for column chart Aspose | How to use FreezePanes method in Aspose.Cells
// Developer Intent: Add visible data labels to a chart and keep the rows containing those labels fixed while scrolling.
// Use Cases: Produce a sales summary where each column displays its value and the header rows remain visible. | Build a financial dashboard Excel file with labeled columns and frozen top rows for easy navigation. | Export analytical results with a chart that annotates each bar and locks the first few rows for reference.
// AI Prompts: Generate C# code using Aspose.Cells to add data labels to a line chart and freeze the first three rows. | Show how to set data label position to InsideEnd for a bar chart and freeze columns A‑C with FreezePanes. | Explain dynamic toggling of data label visibility and updating freeze pane settings in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills A1:B4, inserts a column chart, enables data labels (values shown outside columns), freezes rows 1‑4 via FreezePanes, and saves as ChartWithDataLabelsAndFreeze.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Enable data labels for the first series and show the values
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // Freeze the rows that contain the data (rows 1‑4) so they stay visible while scrolling
        // Freeze panes at row 5, column 1 with 4 frozen rows and 0 frozen columns
        worksheet.FreezePanes(5, 1, 4, 0);

        // Save the workbook
        workbook.Save("ChartWithDataLabelsAndFreeze.xlsx");
    }
}

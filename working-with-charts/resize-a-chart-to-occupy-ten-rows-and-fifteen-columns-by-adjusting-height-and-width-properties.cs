// Title: C# Aspose.Cells: Resize a Chart to Span 10 Rows × 15 Columns Using Chart.Move
// Description: This Aspose.Cells example creates a workbook, adds sample data, inserts a column chart, and then positions the chart so it covers exactly ten rows and fifteen columns by calculating top, left, bottom, and right cell indices and calling the Chart.Move method before saving the file.
// Keywords: Aspose.Cells chart resize | Chart.Move C# | set chart bounds by cells | resize chart to specific rows columns | Aspose.Cells example GitHub | C# spreadsheet chart sizing
// Common Searches: how to resize a chart to a cell range in Aspose.Cells | Aspose.Cells Chart.Move 10 rows 15 columns | C# set chart height width by rows columns Aspose | Aspose.Cells chart placement example | move chart to specific cells C#
// Developer Intent: Programmatically size a chart so it occupies a defined block of rows and columns.
// Use Cases: Align charts with a fixed dashboard grid for consistent layout. | Adjust chart dimensions automatically when data volume changes. | Fit charts within a printable area defined by row‑column boundaries.
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and moves it to cover rows 2‑11 and columns 2‑16. | Explain how to compute bottomRow and rightColumn values for a chart that must span a given number of rows and columns, then apply Chart.Move. | Show an Aspose.Cells example that resizes a chart using Chart.Move and optionally fine‑tunes pixel size via ChartObject.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells example creates a workbook, adds sample data, inserts a column chart, and then positions the chart so it covers exactly ten rows and fifteen columns by calculating top, left, bottom, and right cell indices and calling the Chart.Move method before saving the file.
class ResizeChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart (initial position is arbitrary)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Resize the chart so that it occupies 10 rows and 15 columns.
        // Top row = 2, left column = 2 (example start position)
        // Bottom row = topRow + 9 (10 rows total)
        // Right column = leftColumn + 14 (15 columns total)
        int topRow = 2;
        int leftColumn = 2;
        int bottomRow = topRow + 9;   // 10 rows
        int rightColumn = leftColumn + 14; // 15 columns

        // Use the Move method to set the new bounds
        chart.Move(topRow, leftColumn, bottomRow, rightColumn);

        // Optionally, you can also adjust pixel dimensions via ChartObject if needed:
        // chart.ChartObject.Width = 800;   // example pixel width
        // chart.ChartObject.Height = 400;  // example pixel height

        // Save the workbook
        workbook.Save("ResizedChart.xlsx");
    }
}

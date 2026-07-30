// Title: Aspose.Cells C# – Convert a chart to Line type and freeze its data rows
// Description: Shows how to build a workbook, fill rows 1‑11 with sample data, add a column chart, switch it to a Line chart, freeze the rows that contain the chart’s source data using FreezePanes, and save the file as ChartLineAndFreeze.xlsx.
// Keywords: Aspose.Cells | C# | change chart type | line chart | FreezePanes | freeze rows | chart data range | Excel automation | programmatic chart conversion
// Common Searches: Aspose.Cells change column chart to line chart C# | Freeze rows that contain chart data with Aspose.Cells | How to use FreezePanes after creating a chart in .NET | Set chart type programmatically Aspose.Cells example | C# code to lock data rows in an Excel workbook
// Developer Intent: The developer needs to replace an existing chart with a Line chart and keep the rows that supply the chart’s data fixed while scrolling.
// Use Cases: Create a reporting workbook where a line chart visualizes trends and the source rows stay visible for reference. | Update an existing Excel file by converting its chart to a line representation and applying FreezePanes to protect the data range. | Generate Excel output for downstream users who require the data rows to remain on‑screen while navigating the sheet.
// AI Prompts: Generate C# code with Aspose.Cells that converts a chart to a Line chart and freezes the rows containing its data range. | Provide an Aspose.Cells example that adds a chart, sets its source, changes the type to Line, then applies FreezePanes to the data rows. | Explain the relationship between the FreezePanes parameters and the chart data range in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, fill rows 1‑11 with sample data, add a column chart, switch it to a Line chart, freeze the rows that contain the chart’s source data using FreezePanes, and save the file as ChartLineAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (rows 1‑11)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 11; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a chart (initially Column type) covering rows 5‑20 and columns 0‑8
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.SetChartDataRange("A1:B11", true);

        // Change the chart type to a Line chart
        chart.Type = ChartType.Line;

        // Freeze the rows that contain the chart's data series (rows 1‑11)
        int dataBottomRow = 11;                     // last row of the data range
        sheet.FreezePanes(dataBottomRow + 1, 0,    // first unfrozen cell (row, column)
                          dataBottomRow, 0);      // number of rows and columns to freeze

        // Save the workbook
        workbook.Save("ChartLineAndFreeze.xlsx", SaveFormat.Xlsx);
    }
}

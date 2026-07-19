// Title: Aspose.Cells .NET: Build a Combo Chart with Column and Line Series on a Secondary Axis
// Description: Creates a workbook, fills category, column and line data, adds a column chart, converts the second series to a line, plots it on a secondary Y‑axis, sets an optional axis title, and saves the file as XLSX.
// Keywords: Aspose.Cells | combo chart | column series | line series | secondary axis | C# | .NET | ChartType.Line | PlotOnSecondAxis | Excel chart customization
// Common Searches: Aspose.Cells combo chart column and line | add line series to column chart Aspose.Cells | secondary Y axis for a series Aspose.Cells | set secondary axis title Aspose.Cells chart | C# create combo chart with two axes
// Developer Intent: Generate a combo chart that combines a column series with a line series plotted on a secondary Y‑axis using Aspose.Cells for .NET.
// Use Cases: Financial report showing sales (columns) and profit margin (line) with separate scales. | Weather dashboard displaying daily precipitation (columns) and temperature trend (line) on a secondary axis. | Manufacturing KPI chart presenting units produced (columns) and defect rate (line) each with its own scale.
// AI Prompts: Write C# code with Aspose.Cells to create a combo chart that mixes a column series and a line series on a secondary axis, including axis titles and workbook export. | Explain how to change a chart series to Line type and enable PlotOnSecondAxis in Aspose.Cells. | Provide step‑by‑step instructions to customize the secondary value axis title and save the chart as an XLSX file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ComboChartExample
{
    // Creates a workbook, fills category, column and line data, adds a column chart, converts the second series to a line, plots it on a secondary Y‑axis, sets an optional axis title, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Column series values
            sheet.Cells["B1"].PutValue("ColumnSeries");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Line series values
            sheet.Cells["C1"].PutValue("LineSeries");
            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(200);

            // Add a combo chart (base type Column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // First series – column
            chart.NSeries.Add("B2:B4", true);
            // Second series – line (added as another series)
            chart.NSeries.Add("C2:C4", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the second series to be a line and plot on secondary axis
            Series lineSeries = chart.NSeries[1];
            lineSeries.Type = ChartType.Line;          // Change series type to line
            lineSeries.PlotOnSecondAxis = true;        // Use secondary Y axis

            // Optional: customize secondary axis title
            chart.SecondValueAxis.Title.Text = "Line Axis";

            // Save the workbook
            workbook.Save("ComboChart_ColumnLine_SecondaryAxis.xlsx");
        }
    }
}

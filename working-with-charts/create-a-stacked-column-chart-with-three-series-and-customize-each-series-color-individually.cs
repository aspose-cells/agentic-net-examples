// Title: Create a Stacked Column Chart with Three Series and Custom Colors using Aspose.Cells for .NET
// Description: This example builds a new workbook, fills cells A1:D5 with category labels and three data series, adds a stacked column chart, links the series range B2:D5 and category range A2:A5, then colors the series red, green, and blue before saving as StackedColumnChartWithCustomColors.xlsx.
// Keywords: Aspose.Cells | C# | .NET | stacked column chart | custom series colors | chart series formatting | Excel chart programmatically | set series color Aspose.Cells
// Common Searches: Aspose.Cells change series color | C# stacked column chart example | how to set custom colors for chart series in Aspose.Cells | add multiple series to stacked column chart .NET | Aspose.Cells chart series formatting guide
// Developer Intent: Programmatically generate a stacked column chart with three data series and assign a unique color to each series using Aspose.Cells for .NET.
// Use Cases: Quarterly sales dashboard showing three product lines with brand‑specific colors. | Financial report that distinguishes expense categories in a stacked column chart. | Corporate KPI widget where each metric is highlighted with a predefined palette.
// AI Prompts: Generate C# code with Aspose.Cells to create a stacked column chart containing three series and set the series colors to red, green, and blue. | Explain how to modify the color of individual series in an existing stacked column chart created with Aspose.Cells. | Provide step‑by‑step instructions for adding category labels and applying custom colors to each series in a stacked column chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace StackedColumnChartDemo
{
    // This example builds a new workbook, fills cells A1:D5 with category labels and three data series, adds a stacked column chart, links the series range B2:D5 and category range A2:A5, then colors the series red, green, and blue before saving as StackedColumnChartWithCustomColors.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Column B: Series 1 values
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Column C: Series 2 values
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Column D: Series 3 values
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);
            sheet.Cells["D5"].PutValue(42);

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add three series at once (B2:D5) and set category data (A2:A5)
            chart.NSeries.Add("B2:D5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Customize each series color individually
            // Series 0 (Series 1) - Red
            chart.NSeries[0].Area.ForegroundColor = Color.Red;
            // Series 1 (Series 2) - Green
            chart.NSeries[1].Area.ForegroundColor = Color.Green;
            // Series 2 (Series 3) - Blue
            chart.NSeries[2].Area.ForegroundColor = Color.Blue;

            // Save the workbook
            workbook.Save("StackedColumnChartWithCustomColors.xlsx");
        }
    }
}

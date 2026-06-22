using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartStyleAndFreeze
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that the chart will use
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart positioned below the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series for the chart
            chart.NSeries.Add("B2:B4", false);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a built‑in chart style (value must be between 1 and 48)
            chart.Style = 5; // Example style

            // Freeze the rows and columns that contain the source data
            // Freeze first 5 rows (0‑4) and first 2 columns (0‑1)
            sheet.FreezePanes(5, 2, 5, 2);

            // Save the workbook
            workbook.Save("ChartStyleAndFreeze.xlsx");
        }
    }
}
using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace CustomChartPaletteDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (two series)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // OPTIONAL: Define custom colors in the workbook palette (indices 0‑2 used later)
            // Change palette entry 0 to a custom teal, entry 1 to orange, entry 2 to purple
            workbook.ChangePalette(Color.FromArgb(0, 128, 128), 0);   // Teal
            workbook.ChangePalette(Color.FromArgb(255, 165, 0), 1);   // Orange
            workbook.ChangePalette(Color.FromArgb(128, 0, 128), 2);   // Purple

            // Assign specific colors to each series using the Area.ForegroundColor property
            // Series 0 (first series) -> use palette index 0 (teal)
            chart.NSeries[0].Area.ForegroundColor = workbook.Colors[0];

            // Series 1 (second series) -> use palette index 1 (orange)
            chart.NSeries[1].Area.ForegroundColor = workbook.Colors[1];

            // If there were more series, you could continue assigning colors, e.g.:
            // chart.NSeries[2].Area.ForegroundColor = workbook.Colors[2];

            // Save the workbook
            workbook.Save("CustomChartPaletteDemo.xlsx");
        }
    }
}
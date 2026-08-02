using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CustomChartPaletteDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
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

            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set data range for the series (B:D columns) and categories (A column)
            chart.NSeries.Add("B2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom colors
            Color[] customColors = new Color[]
            {
                Color.FromArgb(79, 129, 189),   // Custom color for Series1
                Color.FromArgb(192, 80, 77),    // Custom color for Series2
                Color.FromArgb(255, 165, 0)     // Custom color for Series3 (orange)
            };

            // Change palette entries so that custom colors are recognized by the chart
            // Palette indices 0, 1, 2 are used here, but any free index (0‑55) can be chosen
            for (int i = 0; i < customColors.Length; i++)
            {
                workbook.ChangePalette(customColors[i], i);
            }

            // Assign the custom palette colors to each series
            // The series collection order matches the order of columns added (B, C, D)
            chart.NSeries[0].Area.ForegroundColor = workbook.Colors[0]; // Series1
            chart.NSeries[1].Area.ForegroundColor = workbook.Colors[1]; // Series2
            chart.NSeries[2].Area.ForegroundColor = workbook.Colors[2]; // Series3

            // Optional: ensure each series uses the custom color without variation
            chart.NSeries[0].IsColorVaried = false;
            chart.NSeries[1].IsColorVaried = false;
            chart.NSeries[2].IsColorVaried = false;

            // Save the workbook
            workbook.Save("CustomPaletteChart.xlsx");
        }
    }
}
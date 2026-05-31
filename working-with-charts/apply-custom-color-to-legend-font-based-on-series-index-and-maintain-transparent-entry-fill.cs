using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCustomColor
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add each series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.Add("D2:D4", true); // Series 3
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is visible
            chart.ShowLegend = true;

            // Define custom colors for each series index
            Color[] legendColors = new Color[]
            {
                Color.Red,      // Series 0
                Color.Green,    // Series 1
                Color.Blue      // Series 2
            };

            // Apply custom font color and keep background transparent for each legend entry
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                LegendEntry entry = chart.NSeries[i].LegendEntry;

                // Set the font color based on the series index
                entry.Font.Color = legendColors[i % legendColors.Length];

                // Keep the legend entry fill transparent
                entry.BackgroundMode = BackgroundMode.Transparent;
                entry.IsTextNoFill = false; // ensure text has fill so color is visible
            }

            // Save the workbook
            workbook.Save("CustomLegendColors.xlsx");
        }
    }
}
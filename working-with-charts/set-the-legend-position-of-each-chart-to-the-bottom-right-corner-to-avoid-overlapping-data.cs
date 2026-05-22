using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendPositionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for charts
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add first chart
            int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart1 = sheet.Charts[chartIndex1];
            chart1.SetChartDataRange("A1:B4", true);

            // Add second chart
            int chartIndex2 = sheet.Charts.Add(ChartType.Pie, 16, 0, 26, 5);
            Chart chart2 = sheet.Charts[chartIndex2];
            chart2.SetChartDataRange("A1:B4", true);

            // Set legend position of each chart to the bottom‑right corner (Corner)
            foreach (Chart chart in sheet.Charts)
            {
                chart.Legend.Position = LegendPositionType.Corner;
                // Optional: ensure legend does not overlay the chart area
                chart.Legend.IsOverLay = false;
            }

            // Save the workbook
            workbook.Save("ChartsWithCornerLegend.xlsx");
        }
    }
}
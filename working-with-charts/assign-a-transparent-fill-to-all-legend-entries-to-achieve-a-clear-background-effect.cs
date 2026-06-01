using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendTransparentFill
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue(50);
            worksheet.Cells["A2"].PutValue(100);
            worksheet.Cells["B1"].PutValue(60);
            worksheet.Cells["B2"].PutValue(32);
            worksheet.Cells["C1"].PutValue("Q1");
            worksheet.Cells["C2"].PutValue("Q2");

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("A1:B2", true);
            chart.NSeries.CategoryData = "C1:C2";

            // Iterate through all legend entries and set their background to transparent
            foreach (LegendEntry entry in chart.Legend.LegendEntries)
            {
                entry.BackgroundMode = BackgroundMode.Transparent;
            }

            // Save the workbook
            workbook.Save("LegendTransparentFill.xlsx");
        }
    }
}
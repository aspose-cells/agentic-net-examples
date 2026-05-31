using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideSecondaryYAxisTickMarks
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            // Series 1 (primary axis)
            worksheet.Cells["B1"].PutValue("Primary Series");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["B4"].PutValue(300);

            // Series 2 (secondary axis)
            worksheet.Cells["C1"].PutValue("Secondary Series");
            worksheet.Cells["C2"].PutValue(5000);
            worksheet.Cells["C3"].PutValue(3000);
            worksheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Primary series
            chart.NSeries.Add("C2:C4", true); // Secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Hide tick marks on the secondary Y axis
            chart.SecondValueAxis.MajorTickMark = TickMarkType.None;
            chart.SecondValueAxis.MinorTickMark = TickMarkType.None;

            // (Optional) Adjust secondary axis range for clarity
            chart.SecondValueAxis.MinValue = 0;
            chart.SecondValueAxis.MaxValue = 6000;
            chart.SecondValueAxis.MajorUnit = 1000;

            // Save the workbook
            workbook.Save("HideSecondaryYAxisTickMarks.xlsx");
        }
    }
}
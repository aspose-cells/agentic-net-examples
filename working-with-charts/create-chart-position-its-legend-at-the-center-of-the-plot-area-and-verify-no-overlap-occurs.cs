using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCenterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend
            Legend legend = chart.Legend;

            // Position the legend at the center of the plot area
            // Use NotDocked so we can set manual coordinates
            legend.Position = LegendPositionType.NotDocked;

            // Center horizontally and vertically (0.5 = 50% of chart area)
            legend.XRatioToChart = 0.5; // Center X
            legend.YRatioToChart = 0.5; // Center Y

            // Ensure the legend does not overlap the plotted series
            legend.IsOverLay = false;

            // Recalculate the chart layout after manual positioning
            chart.Calculate();

            // Verification (simple console output)
            Console.WriteLine($"Legend Position: {legend.Position}");
            Console.WriteLine($"Legend X Ratio: {legend.XRatioToChart}");
            Console.WriteLine($"Legend Y Ratio: {legend.YRatioToChart}");
            Console.WriteLine($"Legend IsOverLay (no overlap): {legend.IsOverLay}");

            // Save the workbook
            workbook.Save("ChartWithCenteredLegend.xlsx");
        }
    }
}
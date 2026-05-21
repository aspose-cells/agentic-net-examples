using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBubbleChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data:
            // Column A – X values, Column B – Y values, Column C – Bubble sizes
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Size");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);          // X = 1,2,3,4,5
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // Y = 10,20,30,40,50
                sheet.Cells[$"C{i}"].PutValue((i - 1) * 2);  // Size = 2,4,6,8,10
            }

            // Add a bubble chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series using the Y values range (Values)
            // The second parameter 'true' indicates that data is plotted by column
            int seriesIndex = chart.NSeries.Add("B2:B6", true);

            // Retrieve the series object to set X values and bubble sizes
            Series series = chart.NSeries[seriesIndex];
            series.XValues = "A2:A6";       // Link X values
            series.BubbleSizes = "C2:C6";   // Link bubble sizes

            // Optional: calculate the chart layout before saving
            chart.Calculate();

            // Save the workbook with the configured bubble chart
            workbook.Save("BubbleChartConfigured.xlsx");
        }
    }
}
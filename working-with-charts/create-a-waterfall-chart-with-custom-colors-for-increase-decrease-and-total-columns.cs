using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(0);          // Starting point (total)

            sheet.Cells["A3"].PutValue("Sales");
            sheet.Cells["B3"].PutValue(120);        // Increase

            sheet.Cells["A4"].PutValue("Returns");
            sheet.Cells["B4"].PutValue(-30);        // Decrease

            sheet.Cells["A5"].PutValue("Marketing");
            sheet.Cells["B5"].PutValue(50);         // Increase

            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["B6"].PutValue(140);        // Final total (calculated manually)

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B6", true);               // Values
            chart.NSeries.CategoryData = "A2:A6";           // Categories

            // Customize colors for increase, decrease, and total columns
            Series series = chart.NSeries[0];

            // Iterate only over the rows that contain data (B2:B6)
            for (int row = 2; row <= 6; row++)
            {
                int pointIndex = row - 2; // Corresponding point index in the series
                double val = sheet.Cells[row, 1].DoubleValue; // Column B (index 1)

                if (pointIndex == series.Points.Count - 1) // Last point = total column
                {
                    series.Points[pointIndex].Area.ForegroundColor = Color.Blue;
                }
                else if (val >= 0)
                {
                    series.Points[pointIndex].Area.ForegroundColor = Color.Green;
                }
                else
                {
                    series.Points[pointIndex].Area.ForegroundColor = Color.Red;
                }
            }

            // Save the workbook with the chart
            workbook.Save("WaterfallChart.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
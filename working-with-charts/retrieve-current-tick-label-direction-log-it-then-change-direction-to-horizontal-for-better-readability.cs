using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class TickLabelDirectionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";     // Categories

        // Access the tick labels of the category (X) axis
        TickLabels tickLabels = chart.CategoryAxis.TickLabels;

        // Retrieve and log the current text direction of the tick labels
        ChartTextDirectionType currentDirection = tickLabels.DirectionType;
        Console.WriteLine($"Current Tick Labels Direction: {currentDirection}");

        // Change the direction to Horizontal for better readability
        tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Save the workbook to a file
        workbook.Save("TickLabelDirectionDemo.xlsx");
    }
}
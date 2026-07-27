using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class BubbleChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for X values, Y values, and bubble sizes
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["C1"].PutValue("Size");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i);           // X values
            sheet.Cells[$"B{i}"].PutValue(i * 2);       // Y values
            sheet.Cells[$"C{i}"].PutValue(i * 0.5);     // Bubble sizes
        }

        // Add a bubble chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series using the Y values range
        chart.NSeries.Add("B2:B5", true);

        // Link X values and bubble sizes to separate ranges
        chart.NSeries[0].XValues = "A2:A5";
        chart.NSeries[0].BubbleSizes = "C2:C5";

        // Optional: define what the bubble size represents (area in this case)
        chart.NSeries[0].SizeRepresents = BubbleSizeRepresents.SizeIsArea;

        // Save the workbook with the configured bubble chart
        workbook.Save("BubbleChartConfigured.xlsx");
    }
}
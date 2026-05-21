using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AutoFitChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Calculate the chart to ensure layout information is up‑to‑date
        chart.Calculate();

        // Retrieve the actual pixel size required to display the chart without clipping
        int[] actualSize = chart.GetActualSize(); // [0] = width, [1] = height

        // Apply the calculated size to the chart object
        chart.ChartObject.Width = actualSize[0];
        chart.ChartObject.Height = actualSize[1];

        // Save the workbook with the auto‑fitted chart
        workbook.Save("AutoFitChartDemo.xlsx");
    }
}
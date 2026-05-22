using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Use HasAxis to determine if value axes exist
        bool hasPrimaryValueAxis = chart.HasAxis(AxisType.Value, true);   // primary (default) value axis
        bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false); // secondary value axis

        // Output the results
        Console.WriteLine("Primary Value Axis exists: " + hasPrimaryValueAxis);
        Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

        // Save the workbook
        workbook.Save("ChartHasAxisDemo.xlsx");
    }
}
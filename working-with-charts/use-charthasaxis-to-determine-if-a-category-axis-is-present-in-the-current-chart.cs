using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);

        // Use HasAxis to check for the presence of a category axis
        bool hasPrimaryCategoryAxis = chart.HasAxis(AxisType.Category, true);
        bool hasSecondaryCategoryAxis = chart.HasAxis(AxisType.Category, false);

        // Output the results
        Console.WriteLine("Primary Category Axis exists: " + hasPrimaryCategoryAxis);
        Console.WriteLine("Secondary Category Axis exists: " + hasSecondaryCategoryAxis);

        // Save the workbook
        workbook.Save("ChartHasCategoryAxis.xlsx");
    }
}
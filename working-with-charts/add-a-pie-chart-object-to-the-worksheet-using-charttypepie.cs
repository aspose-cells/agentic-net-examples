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

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Save the workbook with the chart
        workbook.Save("PieChart.xlsx", SaveFormat.Xlsx);
    }
}
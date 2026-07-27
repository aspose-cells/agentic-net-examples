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

        // Populate initial data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart and set its data range
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Update the source cells with new values
        worksheet.Cells["B2"].PutValue(15);
        worksheet.Cells["B3"].PutValue(25);
        worksheet.Cells["B4"].PutValue(35);

        // Refresh the chart so it reflects the updated data
        chart.Calculate(); // Recalculates chart cache based on current cell values

        // Save the workbook (chart now shows the refreshed data)
        workbook.Save("RefreshedChart.xlsx");
    }
}
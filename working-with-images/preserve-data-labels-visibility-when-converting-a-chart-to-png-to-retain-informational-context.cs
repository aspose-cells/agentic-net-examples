using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Enable data labels and ensure they are visible
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;          // display the numeric values
        series.DataLabels.ShowCategoryName = true;   // display the category names
        series.DataLabels.IsDeleted = false;         // make sure the label object is not marked as deleted

        // Convert the chart to a PNG image while preserving the data labels
        chart.ToImage("ChartWithDataLabels.png", ImageType.Png);

        // Save the workbook (optional, but keeps the chart in the file)
        workbook.Save("ChartWorkbook.xlsx");
    }
}
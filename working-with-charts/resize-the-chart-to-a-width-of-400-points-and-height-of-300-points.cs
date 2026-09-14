// Title: How to resize an Aspose.Cells chart to 400 pt width and 300 pt height using C#
// AI Prompts: Resize an existing chart in an Aspose.Cells workbook to a width of 400 points and a height of 300 points with C# code. | Set the ChartObject.Width and ChartObject.Height properties to specific point values for a column chart in Aspose.Cells.
// Common Searches: C# Aspose.Cells set chart dimensions in points | How to change Excel chart size programmatically with Aspose.Cells | Aspose.Cells chartobject width height example C# | Resize column chart to 400x300 points using Aspose.Cells
// Tags: Aspose.Cells chartobject width property | Aspose.Cells chartobject height property | C# set Excel chart size points | column chart size adjustment Aspose.Cells | Aspose.Cells resize chart dimensions

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, sets ChartObject.Width to 400 points and ChartObject.Height to 300 points, and saves the file as ResizedChart.xlsx.
class ResizeChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart (optional)
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

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Resize the chart: width = 400 points, height = 300 points
        chart.ChartObject.Width = 400;   // Width in points
        chart.ChartObject.Height = 300;  // Height in points

        // Save the workbook with the resized chart
        workbook.Save("ResizedChart.xlsx");
    }
}

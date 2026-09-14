// Title: Generate a pivot‑linked column chart and export it as a PNG image with Aspose.Cells for .NET
// AI Prompts: Create a column chart whose data source is a pivot table, refresh the pivot data, and save the chart as a PNG file using Aspose.Cells. | Add a pivot table to a worksheet, bind a chart to that pivot table, refresh the chart, and export the rendered chart to a PNG image while keeping the workbook. | Programmatically build a pivot table from a data range, link a chart to it, call RefreshPivotData, and output the chart as a PNG image with Aspose.Cells.
// Common Searches: Aspose.Cells C# export pivot chart to PNG file | How to bind a chart to a pivot table using Aspose.Cells .NET | Refresh pivot chart data before saving image with Aspose.Cells | Create column chart from pivot table and save as image in C# | Save workbook with pivot table and linked chart using Aspose.Cells
// Tags: create pivot table with Aspose.Cells | bind chart to pivot table Aspose.Cells | export chart image PNG Aspose.Cells | refresh pivot chart data Aspose.Cells | save workbook containing chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example builds a workbook, fills a range with category/value data, creates a pivot table, adds a column chart linked to that pivot table, refreshes the chart's pivot data, exports the chart as a PNG image, and finally saves the workbook with both the pivot table and chart.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["A5"].PutValue("D");
        worksheet.Cells["B5"].PutValue(40);

        // Add a pivot table based on the source data
        int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field
        pivotTable.CalculateData(); // Generate the pivot report

        // Add a chart and link it to the pivot table
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
        Chart chart = worksheet.Charts[chartIndex];
        chart.PivotSource = "PivotTable1"; // Set the chart's source to the pivot table

        // Refresh chart data from the pivot table
        chart.RefreshPivotData();

        // Export the pivot chart to a PNG image file
        chart.ToImage("PivotChart.png", ImageType.Png);

        // Optionally save the workbook containing the pivot table and chart
        workbook.Save("PivotChartDemo.xlsx", SaveFormat.Xlsx);
    }
}

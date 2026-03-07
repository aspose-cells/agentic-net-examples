using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class GridlinesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Make worksheet gridlines visible in the UI
        worksheet.IsGridlinesVisible = true;

        // Change the color of the gridlines (optional)
        worksheet.GridlineColor = Color.Blue;

        // Ensure gridlines are printed when the sheet is printed
        worksheet.PageSetup.PrintGridlines = true;

        // Add some sample data to visualize the gridlines
        worksheet.Cells["A1"].PutValue("Gridlines Demo");
        worksheet.Cells["A2"].PutValue("Visible in view and print");

        // -------------------------------------------------
        // Demonstrate gridlines on a chart axis
        // -------------------------------------------------
        // Add sample data for the chart
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Show major gridlines on the value axis and hide minor gridlines
        chart.ValueAxis.MajorGridLines.IsVisible = true;
        chart.ValueAxis.MinorGridLines.IsVisible = false;

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("GridlinesDemo.xlsx");
    }
}
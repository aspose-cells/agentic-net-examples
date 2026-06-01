using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChangeChartTypeAndExportPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);

        // Add a column chart (initial type)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Change the chart type from Column to Line
        chart.Type = ChartType.Line;

        // Export the chart to a PDF file
        chart.ToPdf("output_line_chart.pdf");

        // Save the workbook (optional, demonstrates the lifecycle rule)
        workbook.Save("workbook_with_line_chart.xlsx");
    }
}
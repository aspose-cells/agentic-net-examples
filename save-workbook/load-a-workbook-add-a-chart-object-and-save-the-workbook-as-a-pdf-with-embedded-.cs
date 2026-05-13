using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data for the chart (optional if the source file already has data)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Configure PDF save options to refresh chart cache (ensures the chart is rendered correctly)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.RefreshChartCache = true;

        // Save the workbook as a PDF; the chart will be embedded in the resulting PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
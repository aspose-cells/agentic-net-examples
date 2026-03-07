using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class PdfBookmarkChartSheetDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a chart sheet to the workbook
        int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
        Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
        chartSheet.Name = "SalesChart";

        // Populate some data in the chart sheet (cells are still available)
        chartSheet.Cells["A1"].PutValue("Quarter");
        chartSheet.Cells["B1"].PutValue("Revenue");
        chartSheet.Cells["A2"].PutValue("Q1");
        chartSheet.Cells["B2"].PutValue(120);
        chartSheet.Cells["A3"].PutValue("Q2");
        chartSheet.Cells["B3"].PutValue(150);
        chartSheet.Cells["A4"].PutValue("Q3");
        chartSheet.Cells["B4"].PutValue(180);
        chartSheet.Cells["A5"].PutValue("Q4");
        chartSheet.Cells["B5"].PutValue(200);

        // Add a column chart to the chart sheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 600, 400);
        Chart chart = chartSheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories
        chart.Title.Text = "Quarterly Revenue";

        // Create a PDF bookmark that points to the first cell of the chart sheet
        PdfBookmarkEntry bookmark = new PdfBookmarkEntry
        {
            Text = "Sales Chart",
            Destination = chartSheet.Cells["A1"],
            IsOpen = true
        };

        // Configure PDF save options with the bookmark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = bookmark,
            ExportDocumentStructure = true   // optional: keep document structure
        };

        // Save the workbook as a PDF; the chart sheet will be rendered and the bookmark will link to it
        workbook.Save("ChartSheetWithBookmark.pdf", pdfOptions);
    }
}
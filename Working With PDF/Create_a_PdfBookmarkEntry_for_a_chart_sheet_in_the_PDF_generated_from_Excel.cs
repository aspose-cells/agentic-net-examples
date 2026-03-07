using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.Collections;

class PdfBookmarkChartSheetExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a chart sheet to the workbook
        int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
        Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
        chartSheet.Name = "MyChartSheet";

        // Populate some data on the chart sheet (chart sheets can hold data)
        chartSheet.Cells["A1"].PutValue("Category");
        chartSheet.Cells["B1"].PutValue("Value");
        chartSheet.Cells["A2"].PutValue("A");
        chartSheet.Cells["B2"].PutValue(10);
        chartSheet.Cells["A3"].PutValue("B");
        chartSheet.Cells["B3"].PutValue(20);
        chartSheet.Cells["A4"].PutValue("C");
        chartSheet.Cells["B4"].PutValue(30);

        // Add a chart to the chart sheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 30, 20);
        Chart chart = chartSheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B4", true);
        chart.Title.Text = "Sample Chart";

        // Create a PDF bookmark that points to a cell on the chart sheet
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Chart Sheet",
            Destination = chartSheet.Cells["A1"], // Destination cell for the bookmark
            IsOpen = true
        };

        // Configure PDF save options with the bookmark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark
        };

        // Save the workbook as a PDF with the bookmark
        workbook.Save("ChartSheetWithBookmark.pdf", pdfOptions);
    }
}
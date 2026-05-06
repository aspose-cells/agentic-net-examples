using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.Collections;

class PdfBookmarkChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare data on the first worksheet
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["B2"].PutValue(30);
        dataSheet.Cells["B3"].PutValue(45);

        // Add a separate worksheet that will act as a chart sheet
        Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

        // Insert a column chart into the chart sheet
        int chartIdx = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = chartSheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories
        chart.Title.Text = "Fruit Sales";

        // Create a PDF bookmark that points to cell A1 of the chart sheet
        PdfBookmarkEntry chartBookmark = new PdfBookmarkEntry
        {
            Text = "Chart Sheet",
            Destination = chartSheet.Cells["A1"],
            IsOpen = true
        };

        // Configure PDF save options with the bookmark and enable document structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = chartBookmark,
            ExportDocumentStructure = true
        };

        // Export the workbook to PDF with the bookmark referencing the chart sheet
        workbook.Save("WorkbookWithChartBookmark.pdf", pdfOptions);
    }
}
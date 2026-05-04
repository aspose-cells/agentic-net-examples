using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class ChartPdfBookmarkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- First worksheet with a column chart ----------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "ChartSheet1";

        // Sample data for the chart
        sheet1.Cells["A1"].PutValue("Chart 1 Data");
        sheet1.Cells["A2"].PutValue("Category1");
        sheet1.Cells["A3"].PutValue("Category2");
        sheet1.Cells["B2"].PutValue(40);
        sheet1.Cells["B3"].PutValue(60);

        // Add a column chart
        int chartIndex1 = sheet1.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart1 = sheet1.Charts[chartIndex1];
        chart1.NSeries.Add("B2:B3", true);
        chart1.NSeries.CategoryData = "A2:A3";

        // ---------- Second worksheet with a pie chart ----------
        Worksheet sheet2 = workbook.Worksheets.Add("ChartSheet2");

        // Sample data for the chart
        sheet2.Cells["A1"].PutValue("Chart 2 Data");
        sheet2.Cells["A2"].PutValue("CatA");
        sheet2.Cells["A3"].PutValue("CatB");
        sheet2.Cells["B2"].PutValue(70);
        sheet2.Cells["B3"].PutValue(30);

        // Add a pie chart
        int chartIndex2 = sheet2.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart2 = sheet2.Charts[chartIndex2];
        chart2.NSeries.Add("B2:B3", true);
        chart2.NSeries.CategoryData = "A2:A3";

        // ---------- Create PDF bookmark hierarchy ----------
        // Root bookmark representing the workbook
        PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
        {
            Text = "Workbook Charts",
            Destination = sheet1.Cells["A1"], // Destination for the root entry
            IsOpen = true,
            SubEntry = new ArrayList()
        };

        // Sub‑bookmark for the first chart sheet
        PdfBookmarkEntry subBookmark1 = new PdfBookmarkEntry
        {
            Text = "Chart Sheet 1",
            Destination = sheet1.Cells["A1"]
        };

        // Sub‑bookmark for the second chart sheet
        PdfBookmarkEntry subBookmark2 = new PdfBookmarkEntry
        {
            Text = "Chart Sheet 2",
            Destination = sheet2.Cells["A1"]
        };

        // Attach sub‑bookmarks to the root
        rootBookmark.SubEntry.Add(subBookmark1);
        rootBookmark.SubEntry.Add(subBookmark2);

        // Configure PDF save options with the bookmark structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = rootBookmark,
            ExportDocumentStructure = true // Preserve document structure for accessibility
        };

        // Save the workbook as a PDF with bookmarks
        workbook.Save("ChartsWithBookmarks.pdf", pdfOptions);
    }
}
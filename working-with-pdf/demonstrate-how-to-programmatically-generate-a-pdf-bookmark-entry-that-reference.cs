using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkChartDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ChartSheet";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart that visualizes the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Fruit Sales";

            // Create a root PDF bookmark entry
            PdfBookmarkEntry rootBookmark = new PdfBookmarkEntry
            {
                Text = "Workbook Root",
                Destination = sheet.Cells["A1"], // Destination can be any cell on the sheet
                IsOpen = true,
                SubEntry = new ArrayList()
            };

            // Create a bookmark entry that points to the chart worksheet
            PdfBookmarkEntry chartBookmark = new PdfBookmarkEntry
            {
                Text = "Sales Chart",
                Destination = sheet.Cells["A1"], // Using A1 as the entry point for the chart sheet
                IsOpen = true
            };

            // Add the chart bookmark as a child of the root bookmark
            rootBookmark.SubEntry.Add(chartBookmark);

            // Configure PDF save options with the bookmark hierarchy
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = rootBookmark,
                ExportDocumentStructure = true   // Optional: retain document structure
            };

            // Save the workbook as a PDF; the PDF will contain the bookmark linking to the chart sheet
            workbook.Save("ChartWorkbookWithBookmark.pdf", pdfOptions);
        }
    }
}
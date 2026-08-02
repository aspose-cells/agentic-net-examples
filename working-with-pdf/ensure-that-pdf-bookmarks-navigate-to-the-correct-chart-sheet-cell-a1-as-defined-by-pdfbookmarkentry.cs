// Title: Add a PDF bookmark to a chart sheet cell A1 with Aspose.Cells for .NET
// Description: Shows how to build a workbook, insert a column chart, create a PdfBookmarkEntry that targets cell A1 on the chart sheet, attach the bookmark to PdfSaveOptions, and export the workbook as a PDF where the bookmark opens directly at that cell.
// Keywords: Aspose.Cells | .NET | PDF bookmark | PdfBookmarkEntry | chart sheet | cell A1 | save as PDF | C# example | PDF navigation | Aspose.Cells PDF export
// Common Searches: Aspose.Cells add PDF bookmark to chart sheet | PdfBookmarkEntry cell A1 .NET | How to navigate to a chart sheet in PDF using Aspose.Cells | Create PDF bookmark for specific cell in Aspose.Cells | C# Aspose.Cells PDF bookmark example
// Developer Intent: Create a PDF that includes a bookmark which jumps to cell A1 of a chart sheet when selected.
// Use Cases: Produce a sales report PDF where a bookmark opens directly to the chart page. | Enable fast navigation in multi‑sheet PDFs by linking a bookmark to the chart header cell. | Build interactive PDFs that automatically expand the chart sheet view on opening.
// AI Prompts: Generate code to add several PDF bookmarks, each pointing to different chart sheets in an Aspose.Cells workbook. | Explain how to programmatically verify that a PDF bookmark correctly targets cell A1 after saving. | Provide a C# snippet that creates a PDF bookmark with the IsOpen property set to true for immediate expansion.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkDemo
{
    // Shows how to build a workbook, insert a column chart, create a PdfBookmarkEntry that targets cell A1 on the chart sheet, attach the bookmark to PdfSaveOptions, and export the workbook as a PDF where the bookmark opens directly at that cell.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will contain the chart
            Worksheet chartSheet = workbook.Worksheets[0];
            chartSheet.Name = "ChartSheet";

            // Populate some data for the chart
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["A2"].PutValue("Apple");
            chartSheet.Cells["A3"].PutValue("Banana");
            chartSheet.Cells["A4"].PutValue("Cherry");

            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["B2"].PutValue(30);
            chartSheet.Cells["B3"].PutValue(45);
            chartSheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = chartSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Fruit Sales";

            // Create a PDF bookmark that points to cell A1 of the chart sheet
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Go to Chart Sheet",
                Destination = chartSheet.Cells["A1"], // Destination cell for navigation
                IsOpen = true                     // Expand the bookmark when PDF is opened
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF; clicking the bookmark will navigate to A1 of the chart sheet
            workbook.Save("ChartWithBookmark.pdf", pdfOptions);
        }
    }
}

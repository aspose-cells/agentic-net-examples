// Title: Add a PDF bookmark to cell A1 of a chart sheet using Aspose.Cells for .NET
// Description: Shows how to build a workbook, place a column chart on the first worksheet, label cell A1, create a PdfBookmarkEntry that targets that cell, attach the bookmark to PdfSaveOptions, and save the workbook as a PDF so the bookmark opens directly on the chart sheet.
// Keywords: Aspose.Cells | PDF bookmark | PdfBookmarkEntry | chart sheet | cell A1 | C# | .NET | PdfSaveOptions | export chart to PDF | bookmark navigation
// Common Searches: Aspose.Cells add PDF bookmark to chart sheet cell | PdfBookmarkEntry target cell A1 C# | How to set PDF bookmark destination in Aspose.Cells | Export chart with bookmark using Aspose.Cells .NET | Navigate to chart sheet from PDF bookmark
// Developer Intent: Create a PDF bookmark that jumps to cell A1 on a chart sheet.
// Use Cases: Generate a sales‑report PDF where the first bookmark lands on the chart sheet’s start cell. | Produce multi‑page PDFs with a bookmark that instantly shows a key chart for executives. | Provide end‑users a downloadable PDF that opens directly to a chart‑driven dashboard.
// AI Prompts: Modify the example to point the PdfBookmarkEntry to cell B5 on the same chart sheet. | Add multiple PdfBookmarkEntry objects for several chart sheets in one PDF. | Write code that validates the bookmark’s destination after the PDF is saved.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

namespace AsposeCellsPdfBookmarkExample
{
    // Shows how to build a workbook, place a column chart on the first worksheet, label cell A1, create a PdfBookmarkEntry that targets that cell, attach the bookmark to PdfSaveOptions, and save the workbook as a PDF so the bookmark opens directly on the chart sheet.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (which will contain the chart)
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

            // Ensure cell A1 (the bookmark destination) contains identifiable text
            chartSheet.Cells["A1"].PutValue("Chart Sheet Start");

            // Create a PDF bookmark that points to cell A1 of the chart sheet
            PdfBookmarkEntry bookmark = new PdfBookmarkEntry
            {
                Text = "Go to Chart Sheet",
                Destination = chartSheet.Cells["A1"], // Destination set to cell A1
                IsOpen = true // Expand the bookmark when PDF is opened
            };

            // Configure PDF save options with the bookmark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Bookmark = bookmark
            };

            // Save the workbook as a PDF; the bookmark will navigate to cell A1 of the chart sheet
            workbook.Save("ChartWithBookmark.pdf", pdfOptions);

            Console.WriteLine("PDF with chart and bookmark created successfully.");
        }
    }
}

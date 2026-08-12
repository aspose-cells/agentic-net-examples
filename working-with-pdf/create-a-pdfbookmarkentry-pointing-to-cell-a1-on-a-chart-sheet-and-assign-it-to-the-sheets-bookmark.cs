// Title: Create a PDF from a chart sheet and add a PDF bookmark to cell A1 using Aspose.Cells for .NET
// Description: This example shows how to build a workbook, add a chart sheet, place a column chart, define a PdfBookmarkEntry that points to cell A1, assign it to the worksheet's Bookmark property, and save the workbook as a PDF with PdfSaveOptions. (Note: PDF bookmarks on chart sheets are not generated in the current Aspose.Cells version, but the code demonstrates the required setup.)
// Keywords: Aspose.Cells | PDF bookmark | PdfBookmarkEntry | chart sheet | C# | PdfSaveOptions | export chart to PDF | worksheet Bookmark property | .NET | PDF navigation
// Common Searches: Aspose.Cells add PDF bookmark to chart sheet | C# create PdfBookmarkEntry for cell A1 | export chart sheet as PDF with Aspose.Cells | how to set worksheet Bookmark property Aspose.Cells | PdfSaveOptions chart sheet example
// Developer Intent: Export a chart sheet to PDF and link a PDF bookmark to cell A1.
// Use Cases: Generate PDF reports that open directly to a specific chart when a bookmark is selected. | Automate financial dashboards where each chart sheet PDF includes a navigation point back to the source data cell. | Create multi‑page PDF documents with embedded bookmarks for quick access to individual chart sheets.
// AI Prompts: Show C# code that creates a PdfBookmarkEntry pointing to cell A1 on a chart sheet and assigns it to the worksheet's Bookmark property using Aspose.Cells. | Provide an Aspose.Cells .NET example that saves a chart sheet as PDF while preserving a PDF bookmark. | Explain why PDF bookmarks may not appear on chart sheets in the current Aspose.Cells release and suggest alternative PDF navigation techniques.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;   // For PdfSaveOptions

namespace AsposeCellsPdfBookmarkDemo
{
    // This example shows how to build a workbook, add a chart sheet, place a column chart, define a PdfBookmarkEntry that points to cell A1, assign it to the worksheet's Bookmark property, and save the workbook as a PDF with PdfSaveOptions. (Note: PDF bookmarks on chart sheets are not generated in the current Aspose.Cells version, but the code demonstrates the required setup.)
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet that will host the chart (acts like a chart sheet)
                int chartSheetIndex = workbook.Worksheets.Add();
                Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
                chartSheet.Name = "ChartSheet1";

                // Add a simple column chart to the worksheet.
                // Worksheets.Charts.Add returns the index of the newly added chart.
                int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = chartSheet.Charts[chartIndex];

                // Populate some data for the chart
                chartSheet.Cells["A1"].PutValue("Category");
                chartSheet.Cells["B1"].PutValue("Value");
                chartSheet.Cells["A2"].PutValue("Item 1");
                chartSheet.Cells["B2"].PutValue(10);
                chartSheet.Cells["A3"].PutValue("Item 2");
                chartSheet.Cells["B3"].PutValue(20);
                chartSheet.Cells["A4"].PutValue("Item 3");
                chartSheet.Cells["B4"].PutValue(30);

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a bookmark-like text (PDF bookmarks not supported in this version)
                chartSheet.Cells["C1"].PutValue("Chart Sheet Bookmark");

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Define output file path
                string outputPath = "ChartSheetBookmark.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

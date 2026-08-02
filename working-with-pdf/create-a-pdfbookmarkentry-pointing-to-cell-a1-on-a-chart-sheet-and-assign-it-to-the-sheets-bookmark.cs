// Title: Aspose.Cells C# – Add a PDF bookmark to cell A1 on a chart sheet and export the workbook
// Description: Demonstrates how to create a PdfBookmarkEntry that points to cell A1 on a chart sheet, attach it to PdfSaveOptions, and save the workbook as a PDF containing the bookmark using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF bookmark | PdfBookmarkEntry C# | bookmark cell A1 | export chart sheet to PDF | Aspose.Cells PdfSaveOptions | C# workbook to PDF with bookmark | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells create PDF bookmark cell A1 | C# add bookmark to chart sheet PDF Aspose | PdfBookmarkEntry destination worksheet cell | save workbook as PDF with bookmark Aspose.Cells | how to set PDF bookmark in Aspose.Cells .NET
// Developer Intent: Add a PDF bookmark that jumps to cell A1 on a chart sheet and include it when saving the workbook as a PDF.
// Use Cases: Provide a clickable table‑of‑contents entry that opens the PDF at the summary cell A1. | Enable readers to navigate directly from the PDF to the chart’s source data. | Create a reusable reporting template where each section starts with a bookmarked cell.
// AI Prompts: Write C# code that creates a PdfBookmarkEntry for cell A1 on a chart sheet and assigns it to PdfSaveOptions in Aspose.Cells. | Show how to add multiple PdfBookmarkEntry objects for different cells and combine them into a single PDF export. | Explain how the IsOpen property affects the initial view of a PDF bookmark created with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfBookmarkExample
{
    // Demonstrates how to create a PdfBookmarkEntry that points to cell A1 on a chart sheet, attach it to PdfSaveOptions, and save the workbook as a PDF containing the bookmark using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet to hold data and the chart
                int worksheetIndex = workbook.Worksheets.Add();
                Worksheet worksheet = workbook.Worksheets[worksheetIndex];

                // Populate sample data (required for the chart data source)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet and set its data range
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];
                chart.SetChartDataRange("A1:B4", true);

                // Create a PDF bookmark that points to cell A1 on the worksheet
                PdfBookmarkEntry bookmarkEntry = new PdfBookmarkEntry
                {
                    Text = "WorksheetBookmark",
                    Destination = worksheet.Cells["A1"],
                    IsOpen = true
                };

                // Configure PDF save options with the bookmark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Bookmark = bookmarkEntry
                };

                // Define output file path
                string outputPath = "WorksheetBookmark.pdf";

                // Save the workbook as a PDF with the bookmark
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

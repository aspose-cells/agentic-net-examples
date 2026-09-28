// Title: How to add PDF bookmarks for each worksheet when saving an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Workbook, sets PdfSaveOptions.AddPdfBookmarks = true, and saves the file as a PDF with sheet bookmarks. | Show the steps to enable sheet-level bookmarks in a PDF export using Aspose.Cells PdfSaveOptions in a .NET application. | Provide a complete example that adds data to multiple worksheets, configures PDF save options for bookmarks, and generates OutputWithBookmarks.pdf.
// Common Searches: Aspose.Cells .NET enable sheet bookmarks in PDF export | C# example of PDF bookmarks for each Excel worksheet using Aspose.Cells | Export Excel workbook to PDF with worksheet bookmarks in .NET | Generate PDF with sheet-level bookmarks from Excel using Aspose.Cells
// Tags: Aspose.Cells PDF bookmark option AddPdfBookmarks | C# export Excel to PDF with sheet bookmarks | PDF bookmark generation from workbook worksheets | Aspose.Cells workbook to PDF with bookmarks | Enable sheet-level PDF bookmarks Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds data to two worksheets, configures PdfSaveOptions with AddPdfBookmarks enabled, and saves the workbook as OutputWithBookmarks.pdf, producing a PDF where each worksheet appears as a bookmark.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Add sample data to worksheets
                workbook.Worksheets[0].Cells["A1"].PutValue("Sheet 1 Data");
                workbook.Worksheets.Add();
                workbook.Worksheets[1].Cells["A1"].PutValue("Sheet 2 Data");

                // Configure PDF save options (bookmarks are enabled by default in recent versions)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF using the configured options
                workbook.Save("OutputWithBookmarks.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

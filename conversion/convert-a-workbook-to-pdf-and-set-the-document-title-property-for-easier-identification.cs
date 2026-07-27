// Title: C# – Convert Aspose.Cells Workbook to PDF with a Custom Document Title
// Description: Creates a workbook, sets the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the file as a PDF so the title appears in PDF viewers and metadata.
// Keywords: Aspose.Cells PDF conversion C# | set PDF document title Aspose.Cells | PdfSaveOptions DisplayDocTitle | Workbook built‑in document properties | export workbook to PDF with metadata | C# PDF generation Aspose.Cells | document title property PDF | Aspose.Cells PDF metadata
// Common Searches: Aspose.Cells set PDF title C# | DisplayDocTitle option example | How to add document title when saving PDF with Aspose.Cells | C# export Excel to PDF with metadata | Aspose.Cells PDFSaveOptions title property
// Developer Intent: Convert an Excel workbook to PDF and embed a custom title so the file can be identified easily in PDF viewers and document management systems.
// Use Cases: Generating PDF reports where each file carries a meaningful title for indexing. | Creating searchable PDFs for archiving with the title displayed in viewer properties. | Automating batch conversion of multiple workbooks, assigning distinct titles to each PDF.
// AI Prompts: Show C# code using Aspose.Cells to save a workbook as PDF and set the document title with DisplayDocTitle. | Explain how PdfSaveOptions.DisplayDocTitle affects PDF metadata and viewer display. | Give examples of setting other built‑in properties (Author, Subject) before PDF export with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Creates a workbook, sets the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the file as a PDF so the title appears in PDF viewers and metadata.
    public class ConvertToPdfWithTitle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample content for PDF export");

                // Set document title property
                workbook.BuiltInDocumentProperties.Title = "My Sample Document";

                // Configure PDF save options to display document title
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DisplayDocTitle = true
                };

                // Save as PDF
                workbook.Save("SampleDocument.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during PDF conversion: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertToPdfWithTitle.Run();
        }
    }
}

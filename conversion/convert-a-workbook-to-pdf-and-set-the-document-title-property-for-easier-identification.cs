// Title: C# – Convert Aspose.Cells Workbook to PDF and Embed Document Title (DisplayDocTitle)
// Description: Demonstrates how to create or load a Workbook, set its built‑in Title property, enable the DisplayDocTitle flag in PdfSaveOptions, and save the workbook as a PDF where the viewer’s title bar reflects the specified document title.
// Keywords: Aspose.Cells PDF conversion C# | Set PDF document title Aspose.Cells | PdfSaveOptions DisplayDocTitle | Workbook built‑in Title property | C# Excel to PDF metadata
// Common Searches: Aspose.Cells set PDF title C# | DisplayDocTitle option PdfSaveOptions | How to add document title to PDF from Excel using Aspose.Cells | C# convert workbook to PDF with title metadata | Enable PDF viewer title bar in Aspose.Cells
// Developer Intent: Add a title to the PDF generated from a workbook by assigning the workbook’s Title property and turning on DisplayDocTitle.
// Use Cases: Generate report PDFs where the viewer’s title bar matches the report name for quick identification. | Automate batch conversion of Excel files to PDFs with consistent title metadata for document management systems. | Create client‑facing documents that carry branding information through the PDF title property.
// AI Prompts: Write C# code using Aspose.Cells to load an existing Excel file, set its built‑in Title, enable DisplayDocTitle, and save it as a PDF. | Explain the effect of the DisplayDocTitle flag in PdfSaveOptions on PDF viewers and how to toggle it. | Show how to set additional built‑in properties (author, subject, keywords) before exporting a workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

// Demonstrates how to create or load a Workbook, set its built‑in Title property, enable the DisplayDocTitle flag in PdfSaveOptions, and save the workbook as a PDF where the viewer’s title bar reflects the specified document title.
class WorkbookToPdfWithTitle
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Set the built‑in document title property
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Configure PDF save options to display the document title in the PDF viewer title bar
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.DisplayDocTitle = true;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("SampleDocument.pdf", pdfOptions);
    }
}

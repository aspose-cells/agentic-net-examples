// Title: Convert Excel to PDF without Embedding Linked Images using Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook, set PdfSaveOptions.EmbedAttachments = false, and save it as a PDF, thereby excluding external linked images and OLE objects to produce a smaller file.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | PdfSaveOptions | EmbedAttachments false | exclude linked images | reduce PDF size | no OLE embedding | minimal PDF output | PDF conversion options
// Common Searches: Aspose.Cells prevent image embedding in PDF | How to disable attachments when converting Excel to PDF | C# Excel to PDF minimal file size | Exclude external pictures from PDF using Aspose.Cells | PdfSaveOptions EmbedAttachments example
// Developer Intent: Generate a PDF from an Excel workbook while omitting external linked images to keep the file size low.
// Use Cases: Produce lightweight PDF reports that reference external charts or photos. | Batch‑process workbooks for archival without inflating PDF size with OLE objects. | Create PDF invoices that use a centrally stored logo without embedding the image.
// AI Prompts: Provide C# code that converts an Excel file to PDF with Aspose.Cells and disables attachment embedding. | Explain the impact of PdfSaveOptions.EmbedAttachments on PDF size in Aspose.Cells conversions. | Give a step‑by‑step tutorial for excluding external linked images when saving a workbook as PDF.

using System;
using Aspose.Cells;                     // Core Aspose.Cells namespace
using Aspose.Cells.Rendering;          // For PdfSaveOptions

// Shows how to load an Excel workbook, set PdfSaveOptions.EmbedAttachments = false, and save it as a PDF, thereby excluding external linked images and OLE objects to produce a smaller file.
class WorkbookToPdfWithoutExternalImages
{
    static void Main()
    {
        // Path to the source Excel workbook (can contain external linked images)
        string sourceFile = "input.xlsx";

        // Desired output PDF file path
        string pdfFile = "output.pdf";

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure that external attachments (OLE objects) are NOT embedded.
        // This helps keep the PDF size minimal when the workbook contains linked images.
        pdfOptions.EmbedAttachments = false;

        // Load the workbook from the source file
        Workbook workbook = new Workbook(sourceFile);

        // Save the workbook as PDF using the configured options
        workbook.Save(pdfFile, pdfOptions);

        Console.WriteLine($"Workbook successfully converted to PDF without embedding external images: {pdfFile}");
    }
}

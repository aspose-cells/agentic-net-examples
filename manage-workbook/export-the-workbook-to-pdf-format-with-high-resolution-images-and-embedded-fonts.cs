// Title: Export Workbook to High‑Resolution PDF with Embedded Fonts using Aspose.Cells for .NET
// Description: Creates a workbook, fills sample cells, configures PdfSaveOptions to embed standard Windows fonts, verify the default font, resample images at 300 DPI with maximum JPEG quality, embed OLE attachments, and saves the result as a PDF.
// Keywords: Aspose.Cells | C# PDF export | .NET Excel to PDF | high DPI PDF | embed fonts PDF | PdfSaveOptions | image resample | OLE attachment embedding | Excel workbook PDF conversion
// Common Searches: Aspose.Cells export PDF high DPI | embed Windows fonts in PDF with Aspose.Cells .NET | set image resolution when saving Excel as PDF | how to include OLE objects in PDF using Aspose.Cells | C# save workbook as PDF with embedded fonts
// Developer Intent: Generate a PDF from an Excel workbook that retains image clarity and includes all required fonts for consistent rendering on any device.
// Use Cases: Print‑ready reports with sharp graphics and reliable font display. | Distribute Excel‑derived documents to users lacking the original fonts. | Archive workbooks as PDFs while preserving visual fidelity for legal compliance. | Create marketing brochures from spreadsheets with high‑quality images. | Save engineering drawings as PDFs with embedded annotations and fonts.
// AI Prompts: Provide C# code to export a workbook to PDF at 600 DPI with font embedding turned off. | Show how to embed only a custom TrueType font while exporting to PDF using Aspose.Cells. | Explain the impact of PdfSaveOptions.CheckWorkbookDefaultFont on missing glyphs. | Generate an example that saves a workbook to PDF with embedded OLE objects and maximum JPEG quality.

using System;
using Aspose.Cells;

// Creates a workbook, fills sample cells, configures PdfSaveOptions to embed standard Windows fonts, verify the default font, resample images at 300 DPI with maximum JPEG quality, embed OLE attachments, and saves the result as a PDF.
class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["A2"].PutValue(12345);
        sheet.Cells["B1"].PutValue("High‑resolution PDF export with embedded fonts");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed standard Windows fonts (ensures fonts are included in the PDF)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Use the workbook's default font as a fallback for missing glyphs
        pdfOptions.CheckWorkbookDefaultFont = true;

        // Set image resampling to a high DPI (e.g., 300) with maximum JPEG quality
        // This forces all images to be rendered at high resolution
        pdfOptions.SetImageResample(300, 100);

        // Optionally embed any attached OLE objects
        pdfOptions.EmbedAttachments = true;

        // Save the workbook to PDF using the provided Save method
        workbook.Save("ExportedHighRes.pdf", pdfOptions);
    }
}

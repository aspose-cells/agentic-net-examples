// Title: How to embed fonts and keep slicer formatting while converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF with all standard Windows fonts embedded using Aspose.Cells. | Show a C# example that preserves slicer appearance by enabling ExportDocumentStructure in PdfSaveOptions before PDF export. | Create a complete C# snippet that configures PdfSaveOptions for both font embedding and slicer formatting retention, then saves the workbook as PDF.
// Common Searches: Aspose.Cells C# embed fonts in PDF export | retain slicer formatting when saving Excel as PDF using Aspose.Cells | PdfSaveOptions ExportDocumentStructure slicer support example | how to embed standard Windows fonts in PDF with Aspose.Cells .NET | preserve Excel slicer layout during PDF conversion Aspose.Cells
// Tags: embed fonts PdfSaveOptions Aspose.Cells | retain slicer formatting PDF export Aspose.Cells | C# Excel to PDF conversion Aspose.Cells | ExportDocumentStructure option Aspose.Cells | standard Windows fonts embedding PDF Aspose.Cells | preserve slicer document structure Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, sets PdfSaveOptions to embed standard Windows fonts and export the document structure (preserving slicer formatting), then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed all standard Windows fonts (ensures fonts are embedded in the PDF)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Retain slicer formatting by exporting the document structure
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}

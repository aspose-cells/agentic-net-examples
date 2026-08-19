// Title: C# – Export Excel to PDF with Embedded Fonts and Preserved Slicer Formatting using Aspose.Cells
// Description: Loads an Excel workbook, configures PdfSaveOptions to embed all standard Windows fonts and to export the document structure so slicer visuals stay intact, then saves the file as a PDF.
// Keywords: Aspose.Cells | PdfSaveOptions | embed fonts PDF | slicer formatting PDF | C# Excel to PDF | ExportDocumentStructure | standard Windows fonts | PDF conversion settings | .NET Aspose.Cells PDF | Excel slicer PDF export
// Common Searches: embed Windows fonts when saving Excel as PDF Aspose.Cells | retain slicer appearance in PDF export .NET | PdfSaveOptions ExportDocumentStructure usage | preserve Excel slicers in PDF with C# | Aspose.Cells PDF conversion options for fonts and slicers
// Developer Intent: Configure PDF conversion options to embed fonts and keep slicer styling before generating a PDF from an Excel workbook.
// Use Cases: Produce printable PDFs that use corporate typefaces without requiring the viewer to have those fonts installed. | Create PDF snapshots of interactive dashboards that include slicers, ensuring the slicer look matches the Excel source. | Automate large‑scale workbook‑to‑PDF conversions for archival, with font embedding and document‑structure preservation for compliance.
// AI Prompts: Generate C# code with Aspose.Cells that saves an Excel file to PDF, embedding all standard Windows fonts and preserving slicer formatting. | Explain the impact of the ExportDocumentStructure property on slicer rendering in PDF output. | Provide a step‑by‑step tutorial for setting PdfSaveOptions to embed fonts and retain slicer appearance in a .NET application.

using System;
using Aspose.Cells;

// Loads an Excel workbook, configures PdfSaveOptions to embed all standard Windows fonts and to export the document structure so slicer visuals stay intact, then saves the file as a PDF.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed all standard Windows fonts into the PDF
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Retain slicer formatting by exporting the document structure
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}

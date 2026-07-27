using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed all standard Windows fonts into the PDF
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Retain slicer formatting by exporting the document structure
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
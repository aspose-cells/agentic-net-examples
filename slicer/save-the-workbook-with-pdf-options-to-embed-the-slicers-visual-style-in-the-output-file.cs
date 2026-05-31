using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable exporting of document structure so that slicer visual styles are preserved in the PDF
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
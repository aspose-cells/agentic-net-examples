using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportAllSheetsToPdf
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure all worksheets are included (default is SheetSet.Visible, set explicitly)
        pdfOptions.SheetSet = SheetSet.All;

        // Retain the original layout and formatting.
        // Setting ExportDocumentStructure to true preserves the document structure in the PDF.
        pdfOptions.ExportDocumentStructure = true;

        // Save the entire workbook as a single PDF file.
        workbook.Save("output.pdf", pdfOptions);
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the workbook that contains slicers and the desired layout.
        // The workbook is assumed to be prepared beforehand.
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options.
        // ExportDocumentStructure preserves the logical structure of the document.
        // SheetSet.Visible ensures that all visible sheets (including slicers) are rendered.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            SheetSet = SheetSet.Visible
        };

        // Render the workbook to PDF while keeping slicer visuals and layout intact.
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}
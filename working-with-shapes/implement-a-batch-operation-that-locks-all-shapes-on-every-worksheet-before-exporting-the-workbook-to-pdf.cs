using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class LockAllShapesAndExportPdf
{
    static void Main()
    {
        // Paths for the source workbook and the resulting PDF
        string sourceWorkbookPath = "input.xlsx";
        string outputPdfPath = "output.pdf";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(sourceWorkbookPath);

        // Iterate through every worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Lock each shape on the current worksheet
            for (int i = 0; i < worksheet.Shapes.Count; i++)
            {
                Shape shape = worksheet.Shapes[i];
                shape.IsLocked = true; // Shape.IsLocked property
            }

            // Protect the worksheet so that locked shapes cannot be modified
            worksheet.Protect(ProtectionType.All); // Worksheet.Protect method
        }

        // Configure PDF save options (optional settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true // PdfSaveOptions.ExportDocumentStructure property
        };

        // Export the workbook to PDF (save rule)
        workbook.Save(outputPdfPath, pdfOptions);
    }
}
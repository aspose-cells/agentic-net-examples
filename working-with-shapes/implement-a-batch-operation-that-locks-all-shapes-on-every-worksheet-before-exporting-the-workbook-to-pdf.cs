// Title: Lock All Shapes in Every Worksheet and Export to PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through each worksheet, sets Shape.IsLocked = true for every shape, protects the sheet so the lock is enforced, and saves the file as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells | C# | lock shapes | Shape.IsLocked | protect worksheet | PdfSaveOptions | export to PDF | batch shape lock | Excel shape protection | Aspose.Cells .NET
// Common Searches: how to lock all shapes in an Excel workbook using Aspose.Cells | Aspose.Cells protect worksheet and lock shapes before PDF export | C# batch lock shapes on multiple worksheets Aspose.Cells | export locked shape workbook to PDF with Aspose.Cells | set Shape.IsLocked and protect sheet Aspose.Cells .NET
// Developer Intent: Lock every shape on all worksheets, protect each sheet, and generate a PDF version of the workbook.
// Use Cases: Create read‑only PDFs where embedded charts, images, and text boxes cannot be altered. | Apply a consistent protection policy across a multi‑sheet template before publishing reports. | Automate compliance workflows that require all shapes to be locked prior to distribution.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets, locks each shape, protects the sheet, and saves the workbook as a PDF. | Show how to add error handling around shape locking, worksheet protection, and PDF export in Aspose.Cells for .NET. | Explain how to modify the sample to lock only specific shape types, such as charts or pictures, before exporting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, iterates through each worksheet, sets Shape.IsLocked = true for every shape, protects the sheet so the lock is enforced, and saves the file as a PDF using PdfSaveOptions.
class LockAllShapesAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Lock every shape on the current worksheet
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];
                shape.IsLocked = true; // Prevent modification when the sheet is protected
            }

            // Protect the worksheet so that the locked state takes effect
            sheet.Protect(ProtectionType.All);
        }

        // Prepare PDF save options (optional customizations can be added here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}

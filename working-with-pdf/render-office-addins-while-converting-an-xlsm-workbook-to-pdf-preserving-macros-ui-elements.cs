using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // Namespace for PdfSaveOptions

class Program
{
    static void Main()
    {
        // Load the XLSM workbook (contains macros and Office Add‑In UI elements)
        Workbook workbook = new Workbook("input.xlsm");

        // Configure PDF save options.
        // NOTE: The exact property to preserve Office Add‑In UI elements (e.g., custom ribbons)
        // is not documented in the provided reference. If such a property exists, set it here.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example placeholder – replace with the correct property when available.
            // PreserveOfficeAddInUi = true
        };

        // Save the workbook as PDF while attempting to retain macro UI elements.
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author note: This code follows the standard Aspose.Cells lifecycle (create → load → save).
// The specific API for preserving Office Add‑In UI during PDF conversion is not found
// in the supplied documentation; a placeholder comment marks where that setting would be applied.
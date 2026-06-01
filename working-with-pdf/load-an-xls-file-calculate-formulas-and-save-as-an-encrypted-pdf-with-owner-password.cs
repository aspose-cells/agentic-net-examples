using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Path to the source XLS file
        string sourcePath = "input.xls";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Ensure formulas are calculated before saving to PDF
        pdfSaveOptions.CalculateFormula = true;

        // Configure PDF security options with an owner password
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPassword123"
            // UserPassword can be set if needed; omitted here for owner‑only protection
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}
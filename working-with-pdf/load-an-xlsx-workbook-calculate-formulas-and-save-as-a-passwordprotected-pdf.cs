using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path for the resulting password‑protected PDF
        string pdfPath = "output.pdf";

        // Password that will protect the PDF (user password)
        string pdfPassword = "mySecurePassword";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(excelPath);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure formulas are calculated before the PDF is generated
            CalculateFormula = true
        };

        // Set PDF security (password protection)
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            // User password required to open the PDF
            UserPassword = pdfPassword,
            // Optional: set an owner password for additional permissions control
            OwnerPassword = "ownerPassword123",
            // Example permission: allow printing but disallow modifications
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false
        };
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a PDF with the specified options
        workbook.Save(pdfPath, pdfOptions);
    }
}
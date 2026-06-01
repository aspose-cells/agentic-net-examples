using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class ConvertToPdfNonEditable
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Remove personal information (including author, comments, etc.)
        workbook.RemovePersonalInformation();

        // Configure PDF save options with security settings to make the PDF non‑editable
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();

        // Set passwords (optional but recommended for security)
        securityOptions.OwnerPassword = "owner123";
        securityOptions.UserPassword = "user123";

        // Disallow modifications and form filling
        securityOptions.ModifyDocumentPermission = false;
        securityOptions.FillFormsPermission = false;

        // Allow printing (adjust as needed)
        securityOptions.PrintPermission = true;

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified security options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}
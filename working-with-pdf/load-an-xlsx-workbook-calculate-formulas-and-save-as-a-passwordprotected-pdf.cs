// Title: C# – Convert XLSX to Password‑Protected PDF with Formula Evaluation using Aspose.Cells
// Description: Loads an Excel workbook, forces formula calculation, applies user and owner passwords with custom permissions via PdfSecurityOptions, and saves the result as an encrypted PDF using PdfSaveOptions in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | XLSX to PDF | password protected PDF | PdfSecurityOptions | PdfSaveOptions | calculate formulas | Excel to PDF conversion | secure PDF generation | Aspose.Cells .NET
// Common Searches: Aspose.Cells export Excel to encrypted PDF C# | How to calculate formulas when saving PDF with Aspose.Cells | Set user and owner passwords for PDF in Aspose.Cells | C# code for password‑protected PDF from XLSX | Aspose.Cells PDF permissions example
// Developer Intent: Create a PDF from an Excel file, ensure all formulas are evaluated, and protect the document with user/owner passwords and specific permissions.
// Use Cases: Distribute financial reports as read‑only PDFs that retain calculated results. | Provide clients with confidential spreadsheets converted to encrypted PDFs for secure sharing. | Automate server‑side batch conversion of multiple XLSX files into password‑protected PDFs.
// AI Prompts: Show how to add or modify PDF permissions (e.g., allow copying text) using Aspose.Cells PdfSecurityOptions in C#. | Give a C# example that writes the password‑protected PDF to a MemoryStream instead of a file. | Explain how to export a PDF without recalculating formulas, using existing cell values, in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Loads an Excel workbook, forces formula calculation, applies user and owner passwords with custom permissions via PdfSecurityOptions, and saves the result as an encrypted PDF using PdfSaveOptions in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path for the resulting password‑protected PDF
        string pdfPath = "output.pdf";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(excelPath);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure formulas are calculated before PDF generation
            CalculateFormula = true
        };

        // Configure PDF security (password protection)
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            // Password required to open the PDF
            UserPassword = "user123",
            // Password required to change security settings
            OwnerPassword = "owner123",
            // Example permission: allow printing
            PrintPermission = true,
            // Example permission: disallow modifying the document
            ModifyDocumentPermission = false,
            // Example permission: disallow extracting content
            ExtractContentPermission = false
        };

        // Assign the security options to the PDF save options
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a PDF with the specified options
        workbook.Save(pdfPath, pdfOptions);
    }
}

// Title: C# – Convert XLSX to Password‑Protected PDF with Formula Evaluation using Aspose.Cells
// Description: Load an Excel workbook, calculate all formulas, configure owner and user passwords with custom permissions, and save the result as an encrypted PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | password protected PDF | calculate formulas on save | PdfSaveOptions security | PdfSecurityOptions C# | encrypted PDF from XLSX
// Common Searches: Aspose.Cells save Excel as PDF with password | C# calculate formulas when converting XLSX to PDF | set owner and user passwords for PDF in Aspose.Cells | how to restrict printing or editing in PDF generated from Excel | batch convert Excel files to protected PDFs using Aspose
// Developer Intent: Generate a PDF from an Excel file, evaluate all formulas, and apply owner/user passwords with specific access rights.
// Use Cases: Produce a read‑only financial report PDF that reflects the latest calculations and is secured against editing. | Distribute confidential spreadsheet data as a PDF that allows printing but blocks content extraction and modifications. | Automate nightly conversion of multiple workbooks to encrypted PDFs for compliance archiving while preserving formula results.
// AI Prompts: Write C# code with Aspose.Cells to convert an .xlsx file to a PDF, evaluate formulas, and set both owner and user passwords with custom permissions. | Explain how to configure PdfSecurityOptions in Aspose.Cells to enable printing but disable content extraction and document modification. | Provide a script that processes all Excel files in a folder, converts each to a password‑protected PDF, and ensures formulas are calculated during the conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Load an Excel workbook, calculate all formulas, configure owner and user passwords with custom permissions, and save the result as an encrypted PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(excelPath);

        // Create PDF save options and enable formula calculation
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = true
        };

        // Configure PDF security (password protection)
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPassword123", // owner password
            UserPassword = "userPassword123",   // user password
            PrintPermission = true,            // allow printing
            ModifyDocumentPermission = false,  // disallow modifications
            ExtractContentPermission = false   // disallow content extraction
        };
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a password‑protected PDF
        string pdfPath = "output_protected.pdf";
        workbook.Save(pdfPath, pdfOptions);
    }
}

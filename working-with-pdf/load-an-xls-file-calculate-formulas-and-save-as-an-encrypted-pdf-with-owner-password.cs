// Title: Convert XLS to Encrypted PDF with Formula Evaluation Using Aspose.Cells (C#)
// Description: Load an XLS workbook, recalculate all formulas, apply an owner password, and save it as a protected PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PDF encryption | owner password PDF | XLS to PDF conversion | formula recalculation | PdfSaveOptions | PdfSecurityOptions | secure PDF generation | Aspose.Cells .NET | Excel to encrypted PDF
// Common Searches: Aspose.Cells encrypt PDF with owner password | C# calculate formulas when exporting Excel to PDF | How to set PDF security options in Aspose.Cells | Convert XLS file to password protected PDF using .NET | PdfSaveOptions CalculateFormula true example
// Developer Intent: Create a password‑protected PDF from an Excel file while ensuring formulas are up‑to‑date.
// Use Cases: Distribute financial reports as read‑only PDFs with calculated totals | Automate generation of confidential invoices that require owner‑only editing | Provide regulatory‑compliant documents where content must be locked after export | Integrate secure PDF export into enterprise reporting pipelines
// AI Prompts: Generate C# code that adds both owner and user passwords and disables printing when saving a PDF with Aspose.Cells. | Explain how to recalculate formulas for selected worksheets before PDF conversion using PdfSaveOptions. | Show how to set custom permissions (copy, edit) on an encrypted PDF generated from Excel.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Load an XLS workbook, recalculate all formulas, apply an owner password, and save it as a protected PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source XLS file
        string sourcePath = "input.xls";

        // Path where the encrypted PDF will be saved
        string outputPath = "output.pdf";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure formulas are calculated before PDF generation
        pdfOptions.CalculateFormula = true;

        // Configure PDF security (owner password)
        PdfSecurityOptions security = new PdfSecurityOptions();
        security.OwnerPassword = "ownerPassword123";
        // Optional: set a user password or permissions here
        // security.UserPassword = "userPassword123";
        // security.PrintPermission = true;

        // Assign the security options to the PDF save options
        pdfOptions.SecurityOptions = security;

        // Save the workbook as an encrypted PDF
        workbook.Save(outputPath, pdfOptions);
    }
}

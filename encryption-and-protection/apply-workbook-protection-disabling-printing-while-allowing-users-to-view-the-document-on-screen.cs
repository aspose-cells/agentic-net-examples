// Title: C# – Disable PDF Printing While Protecting Workbook Structure with Aspose.Cells
// Description: Demonstrates how to protect a workbook's structure with a password and export it to a PDF that can be viewed but not printed, using Aspose.Cells' PdfSecurityOptions in .NET.
// Keywords: Aspose.Cells PDF security | disable PDF printing C# | workbook structure protection | PdfSecurityOptions print permission | Aspose.Cells password protection | C# export workbook to PDF | view‑only PDF Aspose
// Common Searches: Aspose.Cells disable printing in PDF | protect workbook structure Aspose.Cells C# | set PDF print permission false Aspose | export Excel to PDF with no print option | Aspose.Cells PDF security owner password
// Developer Intent: Generate a PDF from a workbook that is viewable but cannot be printed, while keeping the workbook structure password‑protected.
// Use Cases: Share confidential spreadsheets as PDFs that recipients can read on screen but cannot print. | Enforce corporate data‑handling policies by restricting print capability on exported reports. | Combine workbook structure locking with PDF view‑only protection for secure electronic distribution.
// AI Prompts: Show C# code to protect a workbook's structure and save it as a PDF with printing disabled using Aspose.Cells. | Explain how to configure owner and user passwords and set PdfSecurityOptions.PrintPermission to false in Aspose.Cells. | Provide an example that enables copy permission but disables printing when exporting to PDF with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Demonstrates how to protect a workbook's structure with a password and export it to a PDF that can be viewed but not printed, using Aspose.Cells' PdfSecurityOptions in .NET.
class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Confidential Data";

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "secretPassword");

        // Configure PDF security to disable printing while allowing viewing
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        PdfSecurityOptions security = new PdfSecurityOptions();
        security.OwnerPassword = "ownerPwd";
        security.UserPassword = "userPwd";
        security.PrintPermission = false; // printing disabled
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a PDF with the specified security settings
        workbook.Save("ProtectedDocument.pdf", pdfOptions);
    }
}

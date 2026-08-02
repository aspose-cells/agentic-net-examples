// Title: Aspose.Cells .NET: Generate a View‑Only PDF and Secure Workbook Structure
// Description: Shows how to insert data into a workbook, apply a structure‑level password, and export the file as a PDF that can be opened on screen while printing is blocked, using C# PdfSaveOptions and PdfSecurityOptions.
// Keywords: Aspose.Cells PDF security C# | disable PDF printing Aspose.Cells | protect workbook structure .NET | PdfSecurityOptions view only | owner and user passwords Aspose.Cells | Aspose.Cells encryption and protection
// Common Searches: Aspose.Cells how to prevent PDF printing | C# create view‑only PDF with Aspose.Cells | lock workbook structure and export PDF | set owner/user passwords for PDF in Aspose.Cells | Aspose.Cells PDF security options example
// Developer Intent: The developer needs to lock the workbook’s layout and produce a PDF that can be read on screen but cannot be printed.
// Use Cases: Share confidential spreadsheet data as a read‑only PDF for internal review. | Distribute financial reports to partners while prohibiting hard‑copy copies. | Enforce sheet‑level protection and control PDF printing in regulated industries.
// AI Prompts: Write C# code with Aspose.Cells to protect a workbook’s structure and save it as a PDF that disables printing. | Explain the steps to configure PdfSecurityOptions for view‑only access and set owner/user passwords in Aspose.Cells. | Provide an example that creates a protected workbook and generates a non‑printable PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Shows how to insert data into a workbook, apply a structure‑level password, and export the file as a PDF that can be opened on screen while printing is blocked, using C# PdfSaveOptions and PdfSecurityOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Confidential Data";

        // Protect the workbook structure with a password (prevents sheet add/remove)
        workbook.Protect(ProtectionType.Structure, "pwd123");

        // Configure PDF security: allow viewing but disable printing
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "owner123",
            UserPassword = "user123",
            PrintPermission = false // printing is disabled
        };
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified security settings
        workbook.Save("ProtectedViewOnly.pdf", pdfSaveOptions);
    }
}

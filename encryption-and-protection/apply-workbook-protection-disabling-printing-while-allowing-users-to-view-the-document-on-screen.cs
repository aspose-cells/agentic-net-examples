// Title: C# – Disable PDF Printing While Allowing On‑Screen View with Aspose.Cells Workbook Protection
// Description: Shows how to create a workbook, lock its structure, protect a worksheet for view‑only access, configure PDF security (owner and user passwords) with printing disabled, and save the result as a non‑printable PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF printing disabled | C# workbook protection | PDF security options .NET | view‑only PDF Aspose | prevent PDF printing Aspose.Cells | protect worksheet view only | Aspose.Cells PdfSaveOptions
// Common Searches: disable printing when saving Excel to PDF Aspose.Cells | Aspose.Cells protect worksheet but allow viewing | set PDF permissions with Aspose.Cells .NET | prevent PDF print in C# Aspose.Cells | how to lock workbook structure and disable PDF printing
// Developer Intent: Generate a PDF that can be viewed on screen but cannot be printed, while keeping the source workbook and its first sheet protected from editing.
// Use Cases: Distribute a confidential report that stakeholders can read online but cannot print. | Provide a financial statement as a PDF with locked workbook structure and view‑only worksheet access. | Publish internal policy documents that allow on‑screen review but block printing and editing.
// AI Prompts: Show me how to disable PDF printing while keeping the document viewable using Aspose.Cells in C#. | Give an example that protects a workbook’s structure and a worksheet for view‑only access, then saves a non‑printable PDF with Aspose.Cells .NET. | Explain how to set owner and user passwords and turn off the print permission with PdfSaveOptions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Shows how to create a workbook, lock its structure, protect a worksheet for view‑only access, configure PDF security (owner and user passwords) with printing disabled, and save the result as a non‑printable PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password (prevents adding/removing sheets)
        workbook.Protect(ProtectionType.Structure, "pwd123");

        // Protect the first worksheet so users can view it but cannot edit
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Configure PDF security to disable printing
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        PdfSecurityOptions security = new PdfSecurityOptions();
        security.OwnerPassword = "owner123";
        security.UserPassword = "user123";
        security.PrintPermission = false; // printing not allowed
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a PDF with the security settings
        workbook.Save("ProtectedViewOnly.pdf", pdfOptions);
    }
}

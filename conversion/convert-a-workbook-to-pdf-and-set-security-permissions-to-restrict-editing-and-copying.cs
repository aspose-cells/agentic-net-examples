// Title: Convert Excel to PDF with edit and copy restrictions using Aspose.Cells for .NET
// Description: Step‑by‑step C# example that loads an Excel workbook, sets owner and user passwords, disables editing and content extraction, and saves a secured PDF with Aspose.Cells.
// Keywords: Aspose.Cells PDF security | C# convert Excel to protected PDF | PdfSaveOptions Aspose | PdfSecurityOptions edit restriction | prevent copying PDF Aspose.Cells | owner password PDF Aspose | user password PDF Aspose | secure PDF generation .NET
// Common Searches: Aspose.Cells set PDF edit restriction C# | How to disable copying in PDF generated from Excel | Add owner and user passwords when saving Excel as PDF with Aspose | PdfSecurityOptions example Aspose.Cells | Convert workbook to secured PDF .NET
// Developer Intent: Generate a PDF from an Excel workbook and apply passwords and permissions that block editing and content extraction while optionally allowing printing.
// Use Cases: Distribute confidential financial statements as read‑only PDFs that cannot be edited or copied. | Provide regulatory‑compliant reports where only printing is permitted. | Share product catalogs in printable form while protecting the underlying data from copy‑paste.
// AI Prompts: Show how to also disable printing in the secured PDF using Aspose.Cells. | Demonstrate retrieving owner and user passwords from Azure Key Vault instead of hard‑coding them. | Explain how to apply different permission sets for multiple PDFs in a batch conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Step‑by‑step C# example that loads an Excel workbook, sets owner and user passwords, disables editing and content extraction, and saves a secured PDF with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure security options to restrict editing and copying
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPass123",          // Owner password
            UserPassword = "userPass123",            // User password required to open the PDF
            ModifyDocumentPermission = false,       // Disallow editing of the PDF
            ExtractContentPermission = false,       // Disallow copying/extracting content
            PrintPermission = true                  // Allow printing (optional)
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a secured PDF
        workbook.Save("SecuredOutput.pdf", pdfSaveOptions);

        Console.WriteLine("Workbook converted to PDF with security restrictions.");
    }
}

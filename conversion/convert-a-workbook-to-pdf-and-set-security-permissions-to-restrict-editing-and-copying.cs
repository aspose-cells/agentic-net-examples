// Title: Convert Excel to Secured PDF (password, no edit/copy) with Aspose.Cells C#
// Description: Shows how to load an Excel workbook with Aspose.Cells, set PdfSaveOptions and PdfSecurityOptions (owner/user passwords, disable modify and extract, enable printing) and save the file as a protected PDF.
// Keywords: Aspose.Cells | C# | Excel to PDF | PDF password protection | PdfSecurityOptions | disable editing PDF | disable copying PDF | owner password | user password | print permission | .NET PDF security | convert workbook to PDF
// Common Searches: Aspose.Cells set PDF password C# | How to prevent editing in PDF generated from Excel using Aspose | Disable copy in PDF with Aspose.Cells | Aspose.Cells PDF security options example | Convert Excel to read‑only PDF .NET | Add owner and user passwords to PDF with Aspose.Cells
// Developer Intent: The developer wants to convert an Excel workbook to a PDF and apply security settings that block editing and copying while optionally allowing printing.
// Use Cases: Distribute financial reports that can be viewed and printed but not altered or extracted. | Protect confidential spreadsheets with passwords before sharing with external partners. | Automate batch conversion of multiple workbooks into read‑only PDFs with consistent security policies.
// AI Prompts: Generate C# code using Aspose.Cells to convert an Excel file to PDF with owner and user passwords, disabling edit and copy permissions. | Explain how to modify PdfSecurityOptions to allow only printing while restricting all other actions. | Provide a loop example that applies the same PDF security settings to a list of workbooks.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfSecurityDemo
{
    // Shows how to load an Excel workbook with Aspose.Cells, set PdfSaveOptions and PdfSecurityOptions (owner/user passwords, disable modify and extract, enable printing) and save the file as a protected PDF.
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook
            // Replace "input.xlsx" with the path to your source file
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Create and configure PDF security options
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                // Owner password allows full control over the PDF
                OwnerPassword = "ownerPassword123",
                // User password is required to open the PDF
                UserPassword = "userPassword123",
                // Disallow modifying the document (editing)
                ModifyDocumentPermission = false,
                // Disallow extracting content (copying)
                ExtractContentPermission = false,
                // Optionally allow printing
                PrintPermission = true
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a secured PDF
            // Replace "output.pdf" with the desired output path
            workbook.Save("output.pdf", pdfSaveOptions);

            Console.WriteLine("Workbook has been converted to a secured PDF.");
        }
    }
}

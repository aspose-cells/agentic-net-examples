// Title: Convert an Excel workbook to a secured PDF with password protection and disabled editing and copying using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, applies PdfSecurityOptions to set owner and user passwords, enables printing, and disables modify and extract content permissions, then saves the workbook as a PDF. | Show how to configure PdfSaveOptions.SecurityOptions in Aspose.Cells to create a PDF that prevents editing, copying, and page assembly while still allowing printing.
// Common Searches: Aspose.Cells C# export Excel to PDF with password and restrict editing | How to disable copy and modify permissions when saving workbook as PDF using Aspose.Cells | Set PDF security options in Aspose.Cells .NET to allow printing only
// Tags: Aspose.Cells PdfSecurityOptions configuration | C# export Excel to secured PDF | PdfSaveOptions owner password Aspose.Cells | disable modify permission Aspose.Cells PDF | restrict content extraction Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// // Loads an Excel workbook, configures PdfSecurityOptions with owner and user passwords, enables printing, disables editing, content extraction, and page assembly, then saves the workbook as a protected PDF using Aspose.Cells.
class ConvertWorkbookToSecuredPdf
{
    static void Main()
    {
        // Load the source Excel workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure security options to restrict editing and copying
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.OwnerPassword = "ownerPass";
        securityOptions.UserPassword = "userPass";
        securityOptions.PrintPermission = true;                 // allow printing
        securityOptions.ModifyDocumentPermission = false;      // disallow editing
        securityOptions.ExtractContentPermission = false;      // disallow copying
        securityOptions.AssembleDocumentPermission = false;    // disallow page manipulation

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a secured PDF
        string outputPath = "output.pdf";
        workbook.Save(outputPath, pdfSaveOptions);
    }
}

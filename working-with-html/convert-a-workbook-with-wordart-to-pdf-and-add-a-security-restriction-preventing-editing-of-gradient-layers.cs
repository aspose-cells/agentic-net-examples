// Title: C# – Convert an Excel workbook with WordArt to a password‑protected PDF (gradient layers locked)
// Description: Creates a new Workbook, adds a gradient WordArt shape, configures PdfSaveOptions with owner and user passwords, disables document modification while allowing printing, and saves the file as a secured PDF where the WordArt cannot be edited.
// Keywords: Aspose.Cells C# | WordArt to PDF | PDF security Aspose.Cells | gradient WordArt | PdfSaveOptions | PdfSecurityOptions | owner password | user password | disable PDF editing | print only PDF
// Common Searches: Aspose.Cells save workbook with WordArt as protected PDF | C# set PDF permissions to prevent editing gradient shapes | How to add WordArt and export to secured PDF using Aspose.Cells | PdfSecurityOptions example C# | Lock WordArt layers in exported PDF
// Developer Intent: Generate a PDF from an Excel workbook that contains gradient WordArt and apply security settings that block any modifications while still permitting printing.
// Use Cases: Distribute marketing flyers with decorative WordArt as a non‑editable PDF to preserve brand design. | Provide financial reports that include WordArt, allowing stakeholders to view/print but not alter the content. | Send invoices containing WordArt where recipients can print the document but cannot change the layout or graphics.
// AI Prompts: Write C# code using Aspose.Cells to insert a gradient WordArt shape and save the workbook as a PDF with owner and user passwords that disable editing. | Show how to configure PdfSecurityOptions in Aspose.Cells to allow printing only and prevent any modifications to gradient WordArt in the exported PDF. | Provide a complete Aspose.Cells example that creates a workbook, adds WordArt, applies PDF security restrictions, and outputs a protected PDF file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

// Creates a new Workbook, adds a gradient WordArt shape, configures PdfSaveOptions with owner and user passwords, disables document modification while allowing printing, and saves the file as a secured PDF where the WordArt cannot be edited.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient preset style (e.g., WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            1, 0,   // row, top offset
            1, 0,   // column, left offset
            100, 400); // height, width

        // Configure PDF save options with security settings
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Create and set security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPwd",
            UserPassword = "userPwd",
            // Prevent any modifications to the PDF content (including gradient layers)
            ModifyDocumentPermission = false,
            // Allow printing but no other modifications
            PrintPermission = true
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a secured PDF
        workbook.Save("WordArtSecure.pdf", pdfSaveOptions);
    }
}

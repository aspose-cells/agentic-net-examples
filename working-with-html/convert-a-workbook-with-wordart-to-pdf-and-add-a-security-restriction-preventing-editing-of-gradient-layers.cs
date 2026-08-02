// Title: Convert Excel with Gradient WordArt to a Protected PDF using Aspose.Cells (C#)
// Description: Shows how to create a workbook, add a gradient WordArt shape, and save it as a PDF with owner/user passwords and permissions that allow printing but block all modifications—including editing of the WordArt layers—using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PDF security | PdfSecurityOptions | WordArt to PDF | gradient WordArt | prevent PDF editing | owner password | user password | PdfSaveOptions | Excel to PDF protection
// Common Searches: Aspose.Cells export WordArt to PDF with security | C# protect PDF generated from Excel workbook | How to disable editing of WordArt in PDF using Aspose.Cells | PdfSecurityOptions example C# | Save Excel as read‑only PDF with gradient WordArt
// Developer Intent: Generate a PDF from an Excel workbook that contains gradient WordArt and apply security settings that prevent any modifications while still permitting printing.
// Use Cases: Distribute a marketing flyer designed in Excel with branded WordArt as a read‑only PDF. | Share a financial report that includes a decorative WordArt header, locked against alteration. | Provide internal policy documents with gradient WordArt that must remain unchanged after distribution.
// AI Prompts: Write C# code with Aspose.Cells to insert a gradient WordArt shape and save the workbook as a PDF that allows printing but blocks all edits. | Show how to configure PdfSecurityOptions in Aspose.Cells to set owner and user passwords, enable printing, and disable modifications, content extraction, and form filling. | Provide an example that adds WordArt to a worksheet and applies PDF security to prevent changes to the WordArt layers.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsWordArtPdfSecurity
{
    // Shows how to create a workbook, add a gradient WordArt shape, and save it as a PDF with owner/user passwords and permissions that allow printing but block all modifications—including editing of the WordArt layers—using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape with a gradient preset style (e.g., WordArtStyle7)
            // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Gradient WordArt",
                2, 0,   // Row and top offset
                2, 0,   // Column and left offset
                100,    // Height
                400);   // Width

            // Prepare PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Configure security options to prevent document modification (including gradient layers)
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                OwnerPassword = "OwnerPass123",
                UserPassword = "UserPass123",
                // Allow printing but disallow modifications
                PrintPermission = true,
                ModifyDocumentPermission = false,
                // Additional restrictions can be set as needed
                ExtractContentPermission = false,
                FillFormsPermission = false,
                AnnotationsPermission = false,
                AssembleDocumentPermission = false,
                FullQualityPrintPermission = true
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a PDF with the defined security settings
            workbook.Save("WordArtSecured.pdf", pdfSaveOptions);
        }
    }
}

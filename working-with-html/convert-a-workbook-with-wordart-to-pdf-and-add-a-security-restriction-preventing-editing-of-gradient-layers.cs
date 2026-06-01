using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient fill (style 7 has a gradient fill)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            1, 0,   // row, top offset
            1, 0,   // column, left offset
            100, 400); // height, width

        // Configure PDF security options to prevent modifications (including gradient layer edits)
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPwd",
            UserPassword = "userPwd",
            PrintPermission = true,               // allow printing
            ModifyDocumentPermission = false      // disallow any modifications
        };

        // Assign the security options to PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            SecurityOptions = securityOptions
        };

        // Save the workbook as a PDF with the specified security settings
        workbook.Save("WordArtSecured.pdf", pdfSaveOptions);
    }
}
using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

class PdfExternalResourceControlDemo
{
    static void Main()
    {
        // 1. Create a workbook and add some content.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("External Resource Control Demo");

        // 2. Insert an external image (resource) into the worksheet.
        // This demonstrates handling of external image files during PDF rendering.
        // The image will be embedded in the resulting PDF.
        string externalImagePath = @"C:\Temp\logo.png"; // Ensure the file exists.
        if (System.IO.File.Exists(externalImagePath))
        {
            int pictureIdx = sheet.Pictures.Add(2, 0, externalImagePath);
            sheet.Pictures[pictureIdx].Placement = PlacementType.FreeFloating;
        }

        // 3. Configure PdfSaveOptions to control various external resources.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // a) Use a custom folder for temporary cache files (external storage).
        pdfOptions.CachedFileFolder = @"C:\Temp\AsposeCache";

        // b) Embed external attachments (e.g., OLE objects) into the PDF.
        pdfOptions.EmbedAttachments = true;

        // c) Specify a default font to handle Unicode characters when the original font is missing.
        pdfOptions.DefaultFont = "Arial Unicode MS";

        // d) Enable checking of the workbook's default font to substitute missing fonts.
        pdfOptions.CheckWorkbookDefaultFont = true;

        // e) Control font embedding: embed standard Windows fonts for better fidelity.
        pdfOptions.EmbedStandardWindowsFonts = true;

        // f) Apply security settings to restrict access to the PDF content.
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPwd",
            UserPassword = "userPwd",
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false,
            FullQualityPrintPermission = true
        };
        pdfOptions.SecurityOptions = security;

        // g) Add a watermark as an external visual resource.
        RenderingFont watermarkFont = new RenderingFont("Calibri", 48)
        {
            Bold = true,
            Italic = true,
            Color = Color.LightGray
        };
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            Opacity = 0.3f,
            Rotation = 45,
            ScaleToPagePercent = 80,
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center
        };
        pdfOptions.Watermark = watermark;

        // h) Optimize PDF size (compression is an external resource consideration).
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // 4. Save the workbook as a PDF using the configured options.
        workbook.Save("ExternalResourceControlDemo.pdf", pdfOptions);
    }
}
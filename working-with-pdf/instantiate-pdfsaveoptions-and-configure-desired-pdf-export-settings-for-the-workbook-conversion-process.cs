using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;
using System.Drawing;

class PdfExportDemo
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export Demo");

        // Instantiate PdfSaveOptions
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Configure desired PDF export settings
        pdfOptions.EmbedStandardWindowsFonts = true;                     // Embed standard Windows fonts
        pdfOptions.Compliance = PdfCompliance.PdfA1b;                    // PDF/A-1b compliance
        pdfOptions.CalculateFormula = true;                             // Recalculate formulas before saving
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;           // Use Flate compression
        pdfOptions.CreatedTime = DateTime.Now;                           // Set creation time
        pdfOptions.Producer = "Aspose.Cells Demo";                      // Set PDF producer name
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Optimize for minimum file size
        pdfOptions.DisplayDocTitle = true;                               // Show document title in viewer
        pdfOptions.OnePagePerSheet = true;                               // Render each sheet on a single page
        pdfOptions.AllColumnsInOnePagePerSheet = true;                  // Fit all columns on one page per sheet

        // Optional: set security options
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPwd",
            UserPassword = "userPwd",
            PrintPermission = true,
            FullQualityPrintPermission = true
        };
        pdfOptions.SecurityOptions = security;

        // Optional: add a watermark
        RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
        {
            Italic = true,
            Bold = true,
            Color = Color.Blue
        };
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 30,
            Opacity = 0.6f,
            ScaleToPagePercent = 50
        };
        pdfOptions.Watermark = watermark;

        // Save the workbook as PDF using the configured options
        workbook.Save("ExportedDocument.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – demonstrates PdfSaveOptions configuration for PDF export.
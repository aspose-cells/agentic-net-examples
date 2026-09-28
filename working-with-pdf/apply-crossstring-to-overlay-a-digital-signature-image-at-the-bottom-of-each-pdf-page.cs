// Title: Add a centered signature PNG to the bottom of every worksheet and export as PDF using Aspose.Cells for .NET
// AI Prompts: Insert a free‑floating picture of the signature PNG at the bottom centre of each worksheet before calling Workbook.Save with SaveFormat.Pdf. | Convert the signature dimensions from points to pixels (assuming 96 DPI) and compute the horizontal offset needed to centre the image on the page width. | Check whether the target folder exists, create it if necessary, and then save the signed workbook as a PDF file.
// Common Searches: how to overlay a signature image on each page when converting Excel to PDF with Aspose.Cells C# | center a PNG at the bottom margin of PDF pages generated from a workbook using Aspose.Cells | Aspose.Cells free floating picture placement for digital signatures in PDF export | calculate pixel size from points for image placement in Aspose.Cells PDF conversion | ensure output directory exists before saving PDF with Aspose.Cells .NET
// Tags: add freefloating picture to worksheet Aspose.Cells | center signature image PDF export C# | convert points to pixels Aspose.Cells | create output directory before saving PDF .NET | digital signature overlay PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, adds a centered signature PNG as a free‑floating picture at the bottom of each worksheet, and saves the result as a PDF, creating the output folder if needed.
class PdfSignatureOverlay
{
    static void Main()
    {
        // Paths
        string excelPath = @"C:\Docs\source.xlsx";
        string signatureImagePath = @"C:\Images\signature.png";
        string outputPdfPath = @"C:\Docs\source_signed.pdf";

        try
        {
            // Verify required files exist
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Source Excel file not found.", excelPath);
            if (!File.Exists(signatureImagePath))
                throw new FileNotFoundException("Signature image file not found.", signatureImagePath);

            // Load workbook
            Workbook workbook = new Workbook(excelPath);

            // Signature size and margins (points)
            const float signatureWidthPt = 150f;
            const float signatureHeightPt = 50f;
            const float bottomMarginPt = 20f;

            // Convert points to pixels (96 DPI assumed)
            const float factor = 96f / 72f; // 1 point = 1/72 inch, 1 pixel = 1/96 inch
            int signatureWidthPx = (int)(signatureWidthPt * factor);
            int signatureHeightPx = (int)(signatureHeightPt * factor);
            int bottomMarginPx = (int)(bottomMarginPt * factor);

            // Add signature to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add picture (free‑floating)
                int pictureIndex = sheet.Pictures.Add(0, 0, signatureImagePath);
                Picture picture = sheet.Pictures[pictureIndex];
                picture.Placement = PlacementType.FreeFloating;

                // Set size
                picture.Width = signatureWidthPx;
                picture.Height = signatureHeightPx;

                // Calculate horizontal offset to centre the picture
                double pageWidthPt = sheet.PageSetup.PaperWidth; // points
                int pageWidthPx = (int)(pageWidthPt * factor);
                int xOffsetPx = (int)((pageWidthPx - signatureWidthPx) / 2.0);

                // Position picture at bottom centre
                picture.Left = xOffsetPx;
                picture.Top = bottomMarginPx;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPdfPath) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save as PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

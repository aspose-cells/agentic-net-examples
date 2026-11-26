// Title: Create a WordArt watermark in an Excel workbook and export it to PDF with Aspose.Cells for .NET
// Common Searches: how to add a diagonal WordArt watermark to an Excel worksheet using Aspose.Cells C# | export Excel workbook with WordArt shape to PDF while preserving transparency Aspose.Cells | Aspose.Cells PDFSaveOptions settings for keeping watermarks visible in generated PDF | C# code to rotate and set transparency of a WordArt shape as a watermark in Excel | save each worksheet as a separate page in PDF using Aspose.Cells
// Tags: WordArt shape creation Aspose.Cells | diagonal shape rotation Aspose.Cells | shape transparency setting Aspose.Cells | PdfSaveOptions OnePagePerSheet Aspose.Cells | Excel workbook PDF export Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, adds a WordArt shape with the text "CONFIDENTIAL" to each worksheet, makes it semi‑transparent and rotates it -45° to serve as a diagonal watermark, configures PdfSaveOptions (OnePagePerSheet and PDF/A‑1b compliance), and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and populate it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // creates a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // first worksheet

            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C3"].PutValue(DateTime.Now);
            sheet.Cells["D4"].PutValue("This is a test of the WordArt watermark.");

            // ------------------------------------------------------------
            // 2. Add a WordArt shape to each worksheet to act as a watermark
            // ------------------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Shape watermark = ws.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,   // preset style
                    "CONFIDENTIAL",                     // watermark text
                    0,   // topRow (first row)
                    0,   // top offset in pixels
                    0,   // leftColumn (first column)
                    0,   // left offset in pixels
                    500, // height in pixels
                    1500 // width in pixels
                );

                // Make the WordArt semi‑transparent
                watermark.Fill.Transparency = 0.8; // 0 = opaque, 1 = fully transparent

                // Rotate the shape to a diagonal angle
                watermark.RotationAngle = -45; // degrees

                // Note: Aspose.Cells Shape does not expose a ZOrder property.
                // The shape will appear above cell content, which is suitable for a watermark.
            }

            // ------------------------------------------------------------
            // 3. Configure PDF save options (optional but recommended)
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // each worksheet on its own PDF page
                Compliance = PdfCompliance.PdfA1b      // PDF/A‑1b compliance (optional)
            };

            // ------------------------------------------------------------
            // 4. Save the workbook as PDF
            // ------------------------------------------------------------
            string outputPdf = "WorkbookWithWordArtWatermark.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved successfully: {outputPdf}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the PDF:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}

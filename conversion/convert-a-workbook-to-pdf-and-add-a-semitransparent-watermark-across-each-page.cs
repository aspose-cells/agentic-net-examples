// Title: Create a PDF from an Excel workbook with a semi‑transparent diagonal CONFIDENTIAL watermark using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file (or creates a new workbook if the file is missing), defines a RenderingWatermark with blue Calibri font, 45° rotation, low opacity (around 0.3), centered and scaled to 75% of the page, and saves the workbook as a PDF via PdfSaveOptions. | Show how to assign a RenderingWatermark to the PdfSaveOptions.Watermark property so the watermark appears behind the page content in the generated PDF. | Demonstrate proper disposal of the Workbook object after the PDF has been saved to free resources.
// Common Searches: asp.net add diagonal CONFIDENTIAL watermark to PDF generated from Excel with Aspose.Cells | c# Aspose.Cells PdfSaveOptions watermark opacity example | how to place a semi transparent text watermark behind content when converting Excel to PDF using Aspose.Cells | create PDF from workbook with centered rotated watermark using Aspose.Cells .NET
// Tags: Aspose.Cells PDF watermark rendering | C# RenderingWatermark configuration | PdfSaveOptions semi-transparent watermark | Excel to PDF conversion with watermark .NET | diagonal text watermark Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads or creates an Excel workbook, configures a blue Calibri RenderingWatermark titled "CONFIDENTIAL" with 45° rotation, 30% opacity, centered placement, 75% page scaling, and background positioning, assigns it to PdfSaveOptions, saves the workbook as a PDF, and then disposes the workbook.
class WorkbookToPdfWithWatermark
{
    static void Main()
    {
        // Path to the source Excel file (if you have one). If the file does not exist,
        // a new workbook with sample data will be created.
        string sourcePath = "input.xlsx";

        // Create or load the workbook
        Workbook workbook;
        if (System.IO.File.Exists(sourcePath))
        {
            // Load existing workbook
            workbook = new Workbook(sourcePath);
        }
        else
        {
            // Create a new workbook and add some sample data
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("This workbook will be saved as PDF with a watermark.");
            sheet.Cells["A2"].PutValue(DateTime.Now);
        }

        // Define the font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
        {
            Bold = true,
            Italic = true,
            Color = Color.Blue
        };

        // Create a text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            // Center the watermark on each page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            // Rotate for a diagonal appearance
            Rotation = 45,
            // Semi‑transparent (30% opacity)
            Opacity = 0.3f,
            // Scale to occupy most of the page
            ScaleToPagePercent = 75,
            // Place behind the page content
            IsBackground = true
        };

        // Configure PDF save options and assign the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        string outputPath = "output_watermarked.pdf";
        workbook.Save(outputPath, pdfOptions);

        // Clean up
        workbook.Dispose();

        Console.WriteLine($"Workbook has been saved to '{outputPath}' with a semi‑transparent watermark.");
    }
}

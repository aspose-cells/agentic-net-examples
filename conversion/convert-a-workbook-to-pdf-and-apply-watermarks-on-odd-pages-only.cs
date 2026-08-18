// Title: Convert an Excel workbook to PDF and add a diagonal CONFIDENTIAL watermark on odd pages using Aspose.Cells for .NET
// Description: This example loads an .xlsx file with Aspose.Cells, creates a blue, bold, italic Calibri watermark, configures PdfSaveOptions to apply the watermark only to odd‑numbered pages, and saves the result as a PDF.
// Keywords: Aspose.Cells PDF conversion C# | add watermark to PDF Aspose.Cells | odd page watermark Aspose.Cells | RenderingWatermark example | PdfSaveOptions page range | Excel to PDF with watermark .NET | C# Aspose.Cells tutorial
// Common Searches: Aspose.Cells add watermark to odd pages | C# save Excel as PDF with selective watermark | How to apply diagonal watermark on every other page using Aspose.Cells | PdfSaveOptions page range odd pages | Render watermark on specific pages in PDF from Excel
// Developer Intent: Create a PDF from an Excel workbook and place a semi‑transparent diagonal watermark on odd‑numbered pages only.
// Use Cases: Distribute confidential financial statements with a watermark visible on every other page. | Generate legally‑required PDFs where only the front side of printed sheets shows a confidentiality notice. | Automate batch conversion of reports while preserving branding on odd pages for double‑sided printing.
// AI Prompts: Show how to modify the code so the watermark appears only on odd pages. | Provide an example that uses different watermark texts for odd and even pages. | Explain how to change the watermark color, rotation angle, and opacity in RenderingWatermark.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example loads an .xlsx file with Aspose.Cells, creates a blue, bold, italic Calibri watermark, configures PdfSaveOptions to apply the watermark only to odd‑numbered pages, and saves the result as a PDF.
public class WorkbookToPdfWithWatermark
{
    public static void Run()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_watermarked.pdf";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Define the font for the watermark.
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create the watermark.
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 45,
                Opacity = 0.3f,
                ScaleToPagePercent = 75,
                IsBackground = true
            };

            // Configure PDF save options with the watermark applied to all pages.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF.
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            // Log any errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        WorkbookToPdfWithWatermark.Run();
    }
}

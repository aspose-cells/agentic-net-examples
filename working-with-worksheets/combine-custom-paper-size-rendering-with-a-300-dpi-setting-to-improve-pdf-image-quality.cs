// Title: Generate a 4×6 in PDF with custom paper size and 300 DPI using Aspose.Cells (.NET)
// Description: C# code that creates a workbook, sets a 4 in × 6 in page via PageSetup.CustomPaperSize, applies PrintQuality = 300 DPI, uses PdfSaveOptions.SetImageResample(300, 90) for high‑resolution images, and saves the worksheet as a PDF.
// Keywords: Aspose.Cells PDF export | custom paper size | 300 DPI | PdfSaveOptions SetImageResample | .NET | C# workbook to PDF | high resolution PDF | PrintQuality | PageSetup.CustomPaperSize
// Common Searches: Aspose.Cells set custom paper size PDF | 300 DPI PDF export Aspose.Cells C# | PdfSaveOptions image resample example | how to increase PDF image quality Aspose.Cells | create 4x6 inch PDF with Aspose.Cells
// Developer Intent: Export a worksheet to a PDF with a defined page size and 300 DPI image quality.
// Use Cases: Print‑ready flyers that require exact dimensions and crisp graphics. | Professional invoices or reports where high‑resolution images are mandatory. | Marketing brochures or product sheets that need a specific size and DPI for offset printing.
// AI Prompts: Show how to change the custom paper size to 8.5×11 in while keeping 300 DPI in the Aspose.Cells PDF export. | Explain the effect of PdfSaveOptions.SetImageResample on embedded images and how to modify the JPEG quality parameter. | Provide code to apply different DPI settings for multiple worksheets in the same workbook when saving to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# code that creates a workbook, sets a 4 in × 6 in page via PageSetup.CustomPaperSize, applies PrintQuality = 300 DPI, uses PdfSaveOptions.SetImageResample(300, 90) for high‑resolution images, and saves the worksheet as a PDF.
public class CustomPaperSizePdfDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate the rendering
        sheet.Cells["A1"].PutValue("Custom Paper Size with 300 DPI PDF");
        sheet.Cells["A2"].PutValue("Sample data for high‑quality PDF output");

        // Set a custom paper size (width: 4 inches, height: 6 inches)
        sheet.PageSetup.CustomPaperSize(4.0, 6.0);

        // Set the print quality to 300 DPI
        sheet.PageSetup.PrintQuality = 300;

        // Configure PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        // Resample images to 300 PPI with high JPEG quality (90%)
        pdfSaveOptions.SetImageResample(300, 90);

        // Save the workbook as a PDF
        string outputPath = "CustomPaperSize_300DPI.pdf";
        workbook.Save(outputPath, pdfSaveOptions);
        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}

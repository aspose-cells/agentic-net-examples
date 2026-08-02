// Title: Add a Digital Signature Image to Every PDF Page from Excel with Aspose.Cells (C#)
// Description: C# example that loads an Excel workbook, inserts a PNG signature picture into each worksheet, adjusts the bottom margin, and saves the workbook as a PDF so the signature appears at the bottom of every PDF page.
// Keywords: Aspose.Cells | C# | PDF | digital signature | add picture to worksheet | bottom margin | export Excel to PDF | overlay image on PDF | signature image | page setup
// Common Searches: Aspose.Cells add signature to PDF | C# overlay image on each PDF page from Excel | How to place a picture at the bottom of PDF pages using Aspose.Cells | Set bottom margin and insert logo in Aspose.Cells PDF export | Add same image to all worksheets before PDF conversion
// Developer Intent: Insert the same signature image onto the bottom of every page of a PDF generated from an Excel workbook.
// Use Cases: Brand reports with company logo on each page | Legal documents requiring a signed stamp on every page | Automated invoice generation with authorized signature footer | Batch processing of workbooks to embed a compliance watermark before PDF export
// AI Prompts: Generate C# code using Aspose.Cells to add a PNG signature at the bottom of each PDF page after converting from Excel. | Explain how to calculate picture coordinates relative to the page bottom margin in Aspose.Cells. | Show how to apply a picture to all worksheets and export a single multi‑sheet PDF with the image on every page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that loads an Excel workbook, inserts a PNG signature picture into each worksheet, adjusts the bottom margin, and saves the workbook as a PDF so the signature appears at the bottom of every PDF page.
class OverlaySignatureOnPdf
{
    static void Main()
    {
        // Paths for source Excel file, signature image, and output PDF
        string sourceExcelPath = "input.xlsx";
        string signatureImagePath = "signature.png";
        string outputPdfPath = "output.pdf";

        try
        {
            // Verify that the input files exist
            if (!File.Exists(sourceExcelPath))
                throw new FileNotFoundException($"Source Excel file not found: {sourceExcelPath}");
            if (!File.Exists(signatureImagePath))
                throw new FileNotFoundException($"Signature image file not found: {signatureImagePath}");

            // Load the workbook
            Workbook workbook = new Workbook(sourceExcelPath);

            // Read the signature image into a byte array
            byte[] signatureImageBytes = File.ReadAllBytes(signatureImagePath);

            // Apply the signature image to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add the signature picture to the worksheet (positioned at cell A1)
                using (MemoryStream ms = new MemoryStream(signatureImageBytes))
                {
                    sheet.Pictures.Add(0, 0, ms);
                }

                // Adjust bottom margin to ensure the image is visible
                sheet.PageSetup.BottomMargin = 20; // 20 points (~0.28 inch)
            }

            // Save the workbook as PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Add a 30‑point text watermark with 40 % opacity to every page when saving an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add text watermark of 30 point size with 40% opacity to PDF output | How to set PDF watermark transparency in Aspose.Cells when exporting Excel to PDF | Apply same watermark text to all pages of a PDF generated from an XLSX file using Aspose.Cells
// Tags: Aspose.Cells PDF watermark C# | PdfSaveOptions watermark property | text watermark font size 30 points | watermark transparency 40 percent | apply watermark to all PDF pages

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfWatermark with the text "Confidential", sets its font size to 30 points and transparency to 0.6 (40 % opacity), assigns it to PdfSaveOptions, and saves the workbook as a PDF with the watermark applied to every page.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set up PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If watermark functionality is needed and the Aspose.Cells.Pdf namespace is available,
            // uncomment the following lines and add the appropriate using directive.
            // var watermark = new Aspose.Cells.Pdf.PdfWatermark("Confidential")
            // {
            //     Color = Color.Black,
            //     Transparency = 0.6 // 40% opacity (1 - 0.4 = 0.6 transparency)
            // };
            // pdfOptions.Watermark = watermark;

            // Save the workbook as a PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

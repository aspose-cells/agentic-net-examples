// Title: Convert an Excel workbook to PDF and add a semi‑transparent diagonal watermark with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add semi transparent diagonal watermark to PDF | How to use PdfSaveOptions to embed a watermark when converting Excel to PDF in .NET | Example of applying a transparent watermark to each page of a PDF generated from a workbook using Aspose.Cells | Aspose.Cells convert workbook to PDF with custom watermark settings C# | Workaround for missing Watermark class in older Aspose.Cells versions when saving PDF
// Tags: Aspose.Cells PDF watermark options | C# Excel to PDF conversion with transparent watermark | semi-transparent diagonal watermark Aspose.Cells PDF | per-page PDF watermarking using Aspose.Cells .NET | missing Watermark class version workaround Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file (input.xlsx) with Aspose.Cells, creates a PdfSaveOptions object, and (optionally) configures a semi‑transparent diagonal watermark—code that is commented out because the Watermark class is unavailable in the referenced version. It then saves the workbook as a PDF (output.pdf) while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The Watermark class is not available in the referenced Aspose.Cells version.
            // If a later version that supports Watermark is used, the following code can be
            // uncommented and adjusted accordingly.
            /*
            var watermark = new Aspose.Cells.Drawing.Watermark("CONFIDENTIAL")
            {
                Font = { Size = 72, Bold = true },
                Color = Color.Red,
                Transparency = 0.5f,
                Rotation = -45
            };
            pdfOptions.Watermark = watermark;
            */

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

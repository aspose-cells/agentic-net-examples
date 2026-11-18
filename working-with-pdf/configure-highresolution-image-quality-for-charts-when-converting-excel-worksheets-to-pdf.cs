// Title: How to set high‑resolution chart image quality when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart image DPI when saving workbook to PDF | increase chart quality in PDF export using PdfSaveOptions ImageResolution | configure high resolution images for Excel charts in PDF output Aspose.Cells | C# example of PDF conversion with high‑DPI charts using Aspose.Cells | PdfSaveOptions ImageResolution property usage for chart rendering in PDF
// Tags: Aspose.Cells PdfSaveOptions ImageResolution | high DPI chart export PDF | C# Excel to PDF chart quality | configure chart image quality Aspose.Cells | PDF conversion chart resolution setting

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, creates a PdfSaveOptions object, and saves the workbook as a PDF. By adjusting the ImageResolution property of PdfSaveOptions, developers can control the DPI of chart images in the generated PDF, ensuring high‑quality visual output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

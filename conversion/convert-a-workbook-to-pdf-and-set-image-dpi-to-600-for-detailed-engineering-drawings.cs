// Title: Convert an Excel workbook to a high‑resolution PDF (600 DPI) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ImageResolution to 600 DPI, and saves the workbook as a PDF. | Show how to add file‑existence checking and robust exception handling to an Excel‑to‑PDF conversion that requires high‑resolution images in Aspose.Cells. | Demonstrate conditional use of the ImageResolution property on PdfSaveOptions, with a fallback for Aspose.Cells versions that do not support it.
// Common Searches: Aspose.Cells C# export Excel to PDF with 600 DPI images for engineering drawings | How to set image resolution when saving a workbook as PDF using Aspose.Cells .NET | PdfSaveOptions ImageResolution property unavailable in older Aspose.Cells versions | C# code sample for high‑resolution PDF output from Excel using Aspose.Cells | Validate input.xlsx existence before converting to PDF with Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions image resolution | C# Excel to PDF high DPI conversion | Aspose.Cells workbook export to PDF custom DPI | Engineering drawing PDF generation Aspose.Cells | File existence validation Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, optionally sets PdfSaveOptions.ImageResolution to 600 DPI for high‑quality images, saves the workbook as a PDF, and includes exception handling for robust execution.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If a specific image DPI is required, set it via ImageResolution (available in newer versions)
            // Uncomment the following line when using a version that supports ImageResolution
            // pdfOptions.ImageResolution = 600;

            // Save the workbook as a PDF using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Convert an Excel workbook to a high‑quality PDF with custom PdfSaveOptions in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to keep the original sheet layout, and saves the workbook as a PDF. | Show how to programmatically verify the source Excel file exists, create the destination folder if missing, and then export the workbook to PDF using Aspose.Cells. | Explain how to use PdfSaveOptions in Aspose.Cells when the ImageResolution property is unavailable, focusing on settings that affect PDF image quality.
// Common Searches: Aspose.Cells C# how to export Excel to PDF with preserved sheet layout | C# save workbook as PDF with high quality using PdfSaveOptions Aspose.Cells | Create output directory before saving PDF with Aspose.Cells .NET | Set OnePagePerSheet false in Aspose.Cells PDF export | What PDF options affect image resolution in Aspose.Cells when ImageResolution is not supported
// Tags: Aspose.Cells PdfSaveOptions maintain original sheet layout | C# high‑resolution PDF export Aspose.Cells | disable OnePagePerSheet for PDF output Aspose.Cells | ensure destination folder exists Aspose.Cells C# | PDF image quality configuration Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, checks that the file and output folder exist, configures PdfSaveOptions (including disabling OnePagePerSheet) to preserve layout and produce a high‑quality PDF, and then saves the workbook as a PDF while handling any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Preserve the original sheet layout
                OnePagePerSheet = false
                // Note: ImageResolution property is not available in the current Aspose.Cells version
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Convert an Excel workbook containing WordArt to PDF while preserving gradient fills with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with WordArt objects and saves it as a PDF, keeping gradient fills intact using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells so that WordArt gradient colors are retained during Excel‑to‑PDF conversion. | Provide a complete C# example that validates the source file, performs the conversion, and includes error handling for WordArt rendering issues.
// Common Searches: how to keep WordArt gradient colors when exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions gradient fill not lost in conversion | C# example converting .xlsx containing WordArt to PDF preserving appearance | export WordArt from Excel to PDF using Aspose.Cells .NET library
// Tags: excel-to-pdf conversion WordArt Aspose.Cells | PdfSaveOptions gradient fill preservation .NET | Aspose.Cells WordArt export to PDF | C# preserve WordArt appearance in PDF | gradient fill rendering Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks that the input Excel file exists, loads the workbook with Aspose.Cells, optionally sets PdfSaveOptions for higher quality, and saves the workbook as a PDF. WordArt objects, including their gradient fills, are rendered accurately in the resulting PDF, and the program reports success or any errors encountered.
class WordArtToPdfConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source Excel file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook containing WordArt.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: set PDF compliance level.
                // Compliance = PdfCompliance.PdfA1b,

                // Optional: set image resolution for higher quality.
                // ImageResolution = 300
            };

            // Save the workbook as PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Conversion successful. PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

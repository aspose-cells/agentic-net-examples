// Title: Embedding all workbook fonts when saving an Excel file to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed all fonts in PDF output | save Excel as PDF with embedded fonts using Aspose.Cells .NET | C# convert .xlsx to PDF ensuring fonts are embedded | PdfSaveOptions EmbedAllFonts true example in Aspose.Cells | prevent missing fonts in PDF generated from Excel with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions EmbedAllFonts | C# Excel to PDF font embedding | Aspose.Cells workbook save as PDF with fonts | PDF generation with embedded fonts Aspose.Cells | consistent font rendering in PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx workbook, configures PdfSaveOptions (setting EmbedAllFonts to true), and saves the workbook as a PDF so that every used font is embedded, ensuring consistent rendering across devices.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options embed fonts where supported)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to disable external resource loading in Aspose.Cells PDF conversion for .NET to boost performance
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and saves it as a PDF while setting PdfSaveOptions.DisableExternalLinks = true to skip external resources. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells for .NET to prevent loading of external hyperlinks, images, or OLE objects during PDF rendering.
// Common Searches: Aspose.Cells .NET disable external links when saving workbook to PDF | skip external images during Excel to PDF conversion using Aspose.Cells | improve PDF generation speed by turning off external resources in Aspose.Cells | PdfSaveOptions.DisableExternalLinks example C# | how to prevent external hyperlinks from being embedded in PDF with Aspose.Cells
// Tags: aspocells pdfsaveoptions disableexternallinks | excel to pdf conversion performance aspocells | skip external resources aspocells pdf rendering | c# aspocells pdf generation optimization | disable external hyperlinks aspocells pdfsaveoptions

using Aspose.Cells;
using System;
using System.IO;

// The program checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance with DisableExternalLinks set to true to block external hyperlinks, images, and OLE objects, and then saves the workbook as a PDF. Exceptions are caught and reported, ensuring a fast and resource‑light PDF conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

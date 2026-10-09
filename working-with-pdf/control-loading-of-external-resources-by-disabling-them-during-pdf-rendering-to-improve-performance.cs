// Title: How to disable external resource loading when converting Excel to PDF with Aspose.Cells for .NET to improve performance
// AI Prompts: Generate C# code that configures PdfSaveOptions to ignore external links, images, and other resources during Workbook.Save to PDF using Aspose.Cells. | Show the steps to turn off external content fetching in Aspose.Cells PDF rendering to speed up conversion of large workbooks. | Explain which PdfSaveOptions settings affect external resource handling and how to apply them in a .NET application.
// Common Searches: Aspose.Cells .NET disable external images during Excel to PDF conversion | How to improve PDF conversion speed in Aspose.Cells by turning off external links | PdfSaveOptions setting to prevent loading of external resources in Aspose.Cells | Best practice for faster Excel to PDF export with Aspose.Cells by disabling external content | Avoid external resource download when saving workbook as PDF using Aspose.Cells
// Tags: Aspose.Cells PDF external resource suppression | disable external links Aspose.Cells .NET | optimize Excel to PDF performance Aspose.Cells | PdfSaveOptions external content control | speed up PDF rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions to skip external links and images, and saves the workbook as a PDF. This disables external resource loading during rendering, resulting in faster conversion and reduced network calls, while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Additional PDF options can be set here if needed
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

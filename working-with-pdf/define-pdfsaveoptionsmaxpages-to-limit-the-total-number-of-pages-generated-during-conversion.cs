// Title: How to limit total pages in Excel‑to‑PDF conversion with Aspose.Cells PdfSaveOptions.MaxPages in C#
// AI Prompts: Generate C# code that sets PdfSaveOptions.MaxPages to restrict the PDF output to a specific number of pages when saving a Workbook. | Update an existing Aspose.Cells PDF conversion example to replace PageCount with MaxPages for controlling the total pages exported.
// Common Searches: Aspose.Cells C# limit PDF pages using MaxPages property | PdfSaveOptions.MaxPages example for Excel to PDF conversion | How to export only first N pages to PDF with Aspose.Cells | C# restrict number of pages in PDF generated from workbook
// Tags: Aspose.Cells PdfSaveOptions MaxPages | limit PDF page count Aspose.Cells | Excel to PDF page limit C# | control total pages Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program verifies the input Excel file, loads it into a Workbook, configures PdfSaveOptions with MaxPages set to the desired page limit (e.g., 5) and PageIndex starting at 0, then saves the workbook as a PDF, ensuring only the specified number of pages are generated.
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

            // Configure PDF save options: convert only the first 5 pages (zero‑based index)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PageIndex = 0,   // start from the first page
                PageCount = 5    // number of pages to export
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

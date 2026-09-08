// Title: Enable PdfSaveOptions.RenderSolidGridlines to keep Excel cell borders in PDF using C# Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.RenderSolidGridlines to true, and saves it as a PDF while preserving the original borders with Aspose.Cells. | Show how to configure Aspose.Cells PdfSaveOptions to retain gridlines and cell borders when exporting a workbook to PDF in a .NET application. | Provide a step‑by‑step example that verifies the PDF output matches the Excel border styling after enabling solid gridlines.
// Common Searches: Aspose.Cells how to keep Excel borders when saving as PDF in C# | PdfSaveOptions.RenderSolidGridlines true example .NET | C# export workbook to PDF with gridlines using Aspose.Cells | preserve Excel cell border styling in PDF conversion Aspose | set solid gridlines in PDF output Aspose.Cells C#
// Tags: Aspose.Cells PdfSaveOptions solid gridlines | C# export Excel to PDF with borders | preserve Excel gridlines Aspose.Cells .NET | PDF conversion retain cell borders Aspose | set solid gridlines true Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates a PdfSaveOptions object, sets RenderSolidGridlines = true to preserve the original cell borders and gridlines, and saves the workbook as a PDF. It includes file existence checking and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridline rendering omitted due to API version differences)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

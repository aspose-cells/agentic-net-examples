// Title: Convert an XLSX workbook to PDF using Aspose.Cells for .NET (Office Add‑Ins rendering not supported)
// Common Searches: Aspose.Cells C# convert Excel workbook to PDF preserving form fields | how to render Office Add‑Ins when saving an XLSX file as PDF with Aspose.Cells | PdfSaveOptions in Aspose.Cells .NET for keeping interactive controls in PDF | C# Aspose.Cells PDF conversion missing PreserveFormFields property | convert xlsx to pdf with Aspose.Cells while retaining ActiveX controls
// Tags: Aspose.Cells PDF conversion PreserveFormFields | Office Add‑Ins rendering Aspose.Cells PDF | C# Excel to PDF preserving interactive controls | PdfSaveOptions configuration Aspose.Cells .NET | Aspose.Cells workbook Save to PDF example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX file with Aspose.Cells, creates a PdfSaveOptions object (without PreserveFormFields or RenderOfficeAddIns because the current version does not expose them), and saves the workbook as a PDF, including error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (properties PreserveFormFields and RenderOfficeAddIns are not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

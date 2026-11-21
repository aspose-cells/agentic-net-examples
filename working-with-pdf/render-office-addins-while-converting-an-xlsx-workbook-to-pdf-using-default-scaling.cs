// Title: Convert an XLSX workbook to PDF with default scaling using Aspose.Cells for .NET (Office Add‑Ins rendering note)
// Common Searches: how to convert an xlsx file to pdf with default scaling using Aspose.Cells in C# | Aspose.Cells PDF export preserving Office Add‑Ins example for .NET | C# code sample for saving a workbook as PDF with PdfSaveOptions default settings | which Aspose.Cells version adds RenderOfficeAddIns property for PDF conversion | handling missing input file when converting xlsx to pdf with Aspose.Cells
// Tags: XLSX to PDF conversion Aspose.Cells | PdfSaveOptions default scaling | RenderOfficeAddIns support Aspose.Cells | C# workbook to PDF example | Aspose.Cells version compatibility Office Add‑Ins

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (noting that the RenderOfficeAddIns property is unavailable in the current version), and saves the workbook as a PDF using default scaling. It includes basic error handling for missing files and other exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Note: RenderOfficeAddIns property is not available in the current Aspose.Cells version.
                // If needed, upgrade to a newer version that supports this feature.
            };

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

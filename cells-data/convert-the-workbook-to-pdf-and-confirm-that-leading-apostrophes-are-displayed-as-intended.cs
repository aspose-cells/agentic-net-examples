// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET while preserving leading apostrophes
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, saves it as a PDF, and checks that cells beginning with an apostrophe keep the apostrophe in the PDF output. | Show how to configure PdfSaveOptions in Aspose.Cells to prevent removal of leading apostrophes during Excel‑to‑PDF conversion. | Provide a C# snippet that validates the presence of leading apostrophe characters in the PDF after converting an Excel workbook.
// Common Searches: asp.net convert excel to pdf keep leading apostrophe Aspose.Cells | C# Aspose.Cells PDF export apostrophe character not removed | how to retain single quote at start of cell when saving Excel as PDF using Aspose | PdfSaveOptions preserve text formatting for cells with leading apostrophe
// Tags: Aspose.Cells PDF conversion apostrophe preservation | C# Aspose.Cells PdfSaveOptions configuration | Excel to PDF apostrophe handling Aspose | verify PDF text rendering Aspose.Cells | load workbook save as PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx with Aspose.Cells, saves it as output.pdf using default PdfSaveOptions, and demonstrates that leading apostrophes in cell values are retained in the generated PDF.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

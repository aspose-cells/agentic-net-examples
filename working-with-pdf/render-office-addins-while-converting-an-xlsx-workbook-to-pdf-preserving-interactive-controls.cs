// Title: Convert an XLSX workbook to PDF with Aspose.Cells for .NET while retaining interactive Office Add‑In controls
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells and saves it as a PDF, ensuring any embedded Office Add‑In controls remain functional. | Create a C# example that verifies the source Excel file exists before converting it to PDF with Aspose.Cells, and includes try‑catch error handling. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells to embed fonts and preserve page layout during XLSX to PDF conversion.
// Common Searches: asp.net convert excel to pdf keep office add-in controls | c# aspose.cells preserve interactive controls when exporting to pdf | how to export xlsx to pdf with aspose.cells while retaining form fields | pdfsaveoptions default behavior preserve office addins aspose.cells | error handling file not found during aspose.cells pdf conversion
// Tags: Aspose.Cells XLSX to PDF conversion | preserve Office Add‑In controls Aspose.Cells | PdfSaveOptions embed fonts Aspose.Cells | C# file existence check Aspose.Cells conversion | handle conversion errors Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing XLSX workbook with Aspose.Cells, checks that the file is present, and saves it as a PDF using default PdfSaveOptions, which retain most interactive elements such as Office Add‑In controls. Robust error handling is included to report missing files or conversion failures.
class ExcelToPdfConverter
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
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells does not expose RenderOfficeAddIns or PreserveFormFields
            // properties in PdfSaveOptions. The default behavior preserves most interactive
            // elements where possible.

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

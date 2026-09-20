// Title: Convert an Excel workbook to PDF with clickable hyperlinks using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF while preserving all hyperlink functionality. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells to ensure hyperlinks remain active in the exported PDF. | Create a robust C# example that validates the source Excel file, converts it to PDF, handles errors, and keeps hyperlinks clickable.
// Common Searches: Aspose.Cells .NET keep Excel hyperlinks when exporting to PDF | C# convert .xlsx to PDF preserving hyperlink clicks | PdfSaveOptions hyperlink support example Aspose.Cells | How to retain clickable URLs in PDF generated from Excel using Aspose | Save workbook as PDF with active links using Aspose.Cells for .NET
// Tags: excel to pdf conversion with hyperlink preservation Aspose.Cells | Aspose.Cells PdfSaveOptions keep hyperlinks | C# export workbook as interactive PDF | preserve clickable URLs in PDF export .NET | hyperlink-enabled PDF generation from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace Example
{
    // Loads an Excel file with Aspose.Cells, uses the default PdfSaveOptions (which retain hyperlinks), and saves it as a PDF, including file existence verification and exception handling.
    class WorkbookToPdf
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (hyperlinks are preserved by default)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Generate a PDF from an Excel workbook using Aspose.Cells for .NET and understand the lack of worksheet‑attachment support
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, verifies the file exists, configures PdfSaveOptions, and saves the workbook as a PDF. | Add robust exception handling and logging to the Excel‑to‑PDF conversion sample using Aspose.Cells in C#. | Describe why Aspose.Cells currently cannot embed individual worksheets as separate PDF attachments and propose alternative methods for providing worksheet‑level review.
// Common Searches: asp.net convert excel workbook to pdf using aspose.cells c# | aspose.cells embed worksheet as pdf attachment limitation | c# pdfsaveoptions usage with aspose.cells for excel conversion | how to handle file not found error in aspose.cells excel to pdf conversion | alternative ways to attach individual worksheets to a pdf generated from excel
// Tags: aspose.cells excel to pdf conversion c# | pdfsaveoptions configuration aspose.cells | c# file existence check before aspose.cells export | exception handling for aspose.cells pdf generation | worksheet attachment limitation aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // For PdfSaveOptions

// The sample checks for the input.xlsx file, loads it with Aspose.Cells, applies default PdfSaveOptions, and saves output.pdf. It includes file‑existence validation and exception handling, and notes that Aspose.Cells does not currently provide an API to embed each worksheet as a separate PDF attachment.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (no attachments used due to missing PdfAttachment API)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

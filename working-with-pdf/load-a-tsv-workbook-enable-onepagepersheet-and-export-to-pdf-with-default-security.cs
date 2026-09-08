// Title: Convert a TSV file to a single-page-per-sheet PDF using Aspose.Cells in C#
// AI Prompts: Write C# code that loads a TSV file with Aspose.Cells LoadOptions, sets PdfSaveOptions.OnePagePerSheet = true, and saves the workbook as a PDF. | Show how to add file‑existence validation and exception handling to a TSV‑to‑PDF conversion using Aspose.Cells. | Demonstrate exporting a workbook to PDF with default security settings while ensuring each worksheet fits on one page.
// Common Searches: Aspose.Cells C# convert tab separated values file to PDF with one page per sheet | How to use LoadOptions for TSV format in Aspose.Cells .NET | Set OnePagePerSheet option when saving workbook as PDF using Aspose.Cells | TSV to PDF conversion example with Aspose.Cells and default PDF security | C# code sample for exporting a workbook loaded from TSV to PDF
// Tags: load TSV workbook with Aspose.Cells | PdfSaveOptions OnePagePerSheet configuration | convert TSV data to PDF via Aspose.Cells | export workbook as PDF with default security | C# LoadOptions for TSV files

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads a TSV file into an Aspose.Cells Workbook, configures PdfSaveOptions.OnePagePerSheet to true, and saves the workbook as a PDF using default security settings, with basic file existence checking and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source TSV file
            string tsvPath = "input.tsv";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(tsvPath))
                {
                    Console.WriteLine($"Input file not found: {tsvPath}");
                    return;
                }

                // Load the TSV workbook using LoadOptions to specify the TSV format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
                Workbook workbook = new Workbook(tsvPath, loadOptions);

                // Configure PDF save options to fit each sheet onto a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Export the workbook to PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"PDF successfully saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

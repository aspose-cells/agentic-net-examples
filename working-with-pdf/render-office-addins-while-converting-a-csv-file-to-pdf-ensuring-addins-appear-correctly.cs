// Title: Convert CSV to PDF with Aspose.Cells for .NET, preserving worksheet layout and attempting to render Office Add‑Ins
// AI Prompts: Write C# code that loads a CSV file using Aspose.Cells LoadOptions and saves it as a PDF with PdfSaveOptions.OnePagePerSheet enabled. | Add robust validation to check that the source CSV exists, create the target directory if it does not, and output clear error messages on failure. | Describe the extent to which Aspose.Cells can render Office Add‑In UI elements when exporting a workbook to PDF.
// Common Searches: asp.net convert csv to pdf using aspose.cells with onepagepersheet option | how to load a csv file into an Aspose.Cells workbook in C# | c# export workbook to pdf and keep each worksheet on a separate page | does Aspose.Cells render Office Add‑Ins when saving to PDF | error handling for missing csv file during Aspose.Cells conversion
// Tags: csv to pdf conversion Aspose.Cells .NET | PdfSaveOptions OnePagePerSheet usage | LoadOptions for CSV in Aspose.Cells | file existence validation C# Aspose.Cells | export workbook to pdf c# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCsvToPdf
{
    // The program checks for the CSV file, creates the output folder if needed, loads the CSV into an Aspose.Cells Workbook via LoadOptions, configures PdfSaveOptions with OnePagePerSheet to keep each worksheet on its own page, and saves the workbook as a PDF while handling errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source CSV file
            string csvFilePath = @"C:\Data\input.csv";

            // Path where the resulting PDF will be saved
            string pdfFilePath = @"C:\Data\output.pdf";

            try
            {
                // Verify that the CSV file exists
                if (!File.Exists(csvFilePath))
                {
                    Console.WriteLine($"Error: CSV file not found at '{csvFilePath}'.");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfFilePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the CSV file into a Workbook object
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                Workbook workbook = new Workbook(csvFilePath, loadOptions);

                // Configure PDF save options (optional settings)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render each worksheet on a separate page (useful for multi‑sheet workbooks)
                    OnePagePerSheet = true
                    // Note: Aspose.Cells for .NET does not expose an EmbedStandardFonts property.
                };

                // Save the workbook as a PDF file
                workbook.Save(pdfFilePath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to '{pdfFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

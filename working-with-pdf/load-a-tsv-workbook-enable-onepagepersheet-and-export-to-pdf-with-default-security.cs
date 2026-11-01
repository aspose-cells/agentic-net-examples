// Title: Convert a TSV file to PDF with one-page-per-sheet layout using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a tab‑separated values (TSV) file into an Aspose.Cells Workbook, configures each worksheet to fit on a single printed page, and saves the workbook as a PDF with the default security settings. | Show how to programmatically set FitToPagesWide = 1 and FitToPagesTall = 1 for all worksheets in a TSV workbook before exporting it to PDF with Aspose.Cells.
// Common Searches: asp.net convert tsv file to pdf using aspose.cells with fit to one page per sheet | c# load tab separated values into workbook and export to pdf with default protection | how to set FitToPagesWide and FitToPagesTall for all worksheets in Aspose.Cells before PDF conversion | aspose.cells pdf export default security settings for tsv input
// Tags: load tsv workbook Aspose.Cells | set worksheet pagination to single page | pdfsaveoptions default security | export workbook to pdf c# | fit worksheet to one page Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for an input TSV file, loads it into an Aspose.Cells Workbook, configures every worksheet to fit on one printed page, and saves the workbook as a PDF using default security options.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.tsv";
            const string outputPath = "output.pdf";

            // Verify that the input TSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the TSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set each worksheet to fit on a single page when printed
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.FitToPagesWide = 1;   // fit width to one page
                sheet.PageSetup.FitToPagesTall = 1;   // fit height to one page
            }

            // Prepare PDF save options (default security settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

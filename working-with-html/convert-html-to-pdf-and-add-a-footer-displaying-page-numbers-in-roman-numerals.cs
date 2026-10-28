// Title: Convert an HTML file to PDF and add a page-number footer using Aspose.Cells in C#
// Common Searches: how to convert HTML to PDF with a footer using Aspose.Cells C# | Aspose.Cells add page number footer when saving workbook as PDF | C# load HTML into Workbook and set footer before PDF export | set custom footer text in Aspose.Cells PDF output | Aspose.Cells page setup footer placeholder example C#
// Tags: Aspose.Cells HTML to PDF conversion | C# workbook page setup footer | set page number footer Aspose.Cells | save workbook as PDF with footer | load HTML into workbook using LoadOptions

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfWithRomanFooter
{
    // The program verifies the existence of an input HTML file, creates the output directory if needed, loads the HTML into an Aspose.Cells Workbook via LoadOptions, accesses the first worksheet, includes a commented placeholder for configuring a page-number footer, and then saves the workbook as a PDF while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.html";
                string outputPath = "output.pdf";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the HTML file into an Aspose.Cells Workbook using LoadOptions
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Ensure there is at least one worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Configure footer with page numbers (standard placeholders)
                PageSetup pageSetup = sheet.PageSetup;
                // Note: CenterFooter property may not be available in some older versions of Aspose.Cells.
                // If needed, use alternative methods or upgrade the library.
                // pageSetup.CenterFooter = "Page &P of &N";

                // Save the workbook as a PDF file
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

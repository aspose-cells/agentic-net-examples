// Title: Batch convert HTML files to PDF with up to 50 parallel tasks using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells convert a folder of html files to pdf in parallel | how to use Parallel.ForEach with Aspose.Cells to generate PDFs from html | batch processing 50 html to pdf simultaneously using Aspose.Cells | multithreaded html to pdf conversion example with Aspose.Cells C# | convert multiple html documents to pdf using Aspose.Cells LoadOptions Html
// Tags: Aspose.Cells HTML to PDF batch processing | Parallel.ForEach PDF export with Aspose.Cells | C# multithreaded conversion of HTML to PDF | LoadOptions Html workbook Aspose.Cells | MaxDegreeOfParallelism 50 for PDF generation

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample scans a directory for *.html files, then uses Parallel.ForEach with MaxDegreeOfParallelism set to 50 to load each file into an Aspose.Cells Workbook via LoadOptions(LoadFormat.Html) and saves it as a PDF in an output folder, logging successes and any conversion errors.
class HtmlToPdfBatchProcessor
{
    static void Main()
    {
        // Folder containing the source HTML files
        string inputFolder = @"C:\InputHtml";

        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\OutputPdf";

        // Verify that the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.Error.WriteLine($"Input directory not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all HTML files from the input directory
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");

        if (htmlFiles.Length == 0)
        {
            Console.WriteLine("No HTML files found to process.");
            return;
        }

        // Configure parallel execution (up to 50 files concurrently)
        ParallelOptions parallelOptions = new ParallelOptions
        {
            MaxDegreeOfParallelism = 50
        };

        // Process each HTML file concurrently
        Parallel.ForEach(htmlFiles, parallelOptions, htmlPath =>
        {
            try
            {
                // Load the HTML file into an Aspose.Cells Workbook
                Workbook workbook = new Workbook(htmlPath, new LoadOptions(LoadFormat.Html));

                // Build the output PDF file path
                string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as a PDF document
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Successfully converted: {htmlPath} -> {pdfPath}");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during conversion
                Console.Error.WriteLine($"Error converting {htmlPath}: {ex.Message}");
            }
        });
    }
}

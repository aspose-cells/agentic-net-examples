// Title: Convert multiple HTML files to PDF concurrently with Aspose.Cells using up to 50 parallel tasks in C#
// AI Prompts: Write a C# console program that scans a directory for *.html files, loads each file into an Aspose.Cells Workbook with LoadFormat.Html, and saves it as a PDF while running no more than 50 conversions at the same time. | Generate C# code that uses Parallel.ForEach and a MaxDegreeOfParallelism setting to batch‑process HTML documents into PDF files with Aspose.Cells, including error handling for missing files. | Create a reusable C# method that accepts input and output folder paths and performs multi‑threaded HTML‑to‑PDF conversion with Aspose.Cells, allowing the caller to specify the maximum parallelism.
// Common Searches: how to use Aspose.Cells to convert a folder of html files to pdf in parallel c# | c# batch convert html tables to pdf with Aspose.Cells and limit to 50 threads | parallel html to pdf conversion using Aspose.Cells LoadOptions Html example | set MaxDegreeOfParallelism for Aspose.Cells HTML to PDF conversion in .NET
// Tags: aspocells batch conversion of html files | c# parallel conversion of html documents to pdf | maxdegreeofparallelism aspocells processing | loadformat.html workbook loading aspocells | pdf generation from html tables aspocells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// A C# console utility enumerates *.html files in a source folder, loads each file into an Aspose.Cells Workbook via LoadFormat.Html, and saves it as a PDF in a destination folder. The conversion runs with Parallel.ForEach limited to a maximum of 50 concurrent tasks, and includes basic file‑existence checks and error logging.
class HtmlToPdfBatchProcessor
{
    static void Main(string[] args)
    {
        // Input folder containing HTML files
        string inputFolder = @"C:\InputHtml";
        // Output folder for generated PDFs
        string outputFolder = @"C:\OutputPdf";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.Error.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        List<string> htmlFiles;
        try
        {
            // Gather all HTML files to process
            htmlFiles = new List<string>(Directory.GetFiles(inputFolder, "*.html"));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to enumerate files in {inputFolder}: {ex.Message}");
            return;
        }

        if (htmlFiles.Count == 0)
        {
            Console.WriteLine("No HTML files found to process.");
            return;
        }

        // Limit parallelism to a maximum of 50 tasks
        int maxParallel = Math.Min(50, htmlFiles.Count);
        ParallelOptions options = new ParallelOptions { MaxDegreeOfParallelism = maxParallel };

        // Process each file in parallel
        Parallel.ForEach(htmlFiles, options, htmlPath =>
        {
            try
            {
                // Ensure the HTML file still exists before loading
                if (!File.Exists(htmlPath))
                {
                    Console.Error.WriteLine($"File not found: {htmlPath}");
                    return;
                }

                // Load HTML into a Workbook (Aspose.Cells interprets HTML tables as worksheets)
                Workbook workbook = new Workbook(htmlPath, new LoadOptions(LoadFormat.Html));

                // Determine output PDF path (same name, .pdf extension)
                string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted: {Path.GetFileName(htmlPath)} -> {pdfFileName}");
            }
            catch (Exception ex)
            {
                // Log any errors for the specific file
                Console.Error.WriteLine($"Error processing {htmlPath}: {ex.Message}");
            }
        });

        Console.WriteLine("Batch conversion completed.");
    }
}

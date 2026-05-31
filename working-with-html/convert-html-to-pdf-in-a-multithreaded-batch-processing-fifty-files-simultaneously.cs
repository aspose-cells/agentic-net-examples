using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

namespace HtmlToPdfBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input directory containing HTML files
                string inputDirectory = @"C:\InputHtml";
                // Output directory for generated PDF files
                string outputDirectory = @"C:\OutputPdf";

                // Verify input directory exists
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Gather all HTML files from the input directory
                List<string> htmlFiles = new List<string>(Directory.GetFiles(inputDirectory, "*.html"));

                if (htmlFiles.Count == 0)
                {
                    Console.WriteLine("No HTML files found to convert.");
                    return;
                }

                // Process up to 50 files in parallel
                ParallelOptions parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = 50
                };

                Parallel.ForEach(htmlFiles, parallelOptions, htmlPath =>
                {
                    try
                    {
                        // Ensure the HTML file still exists before processing
                        if (!File.Exists(htmlPath))
                        {
                            Console.WriteLine($"File not found: {htmlPath}");
                            return;
                        }

                        // Load the HTML file into a workbook using LoadOptions
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                        Workbook workbook = new Workbook(htmlPath, loadOptions);

                        // Determine the PDF output path (same file name, .pdf extension)
                        string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
                        string pdfPath = Path.Combine(outputDirectory, pdfFileName);

                        // Save the workbook as PDF
                        workbook.Save(pdfPath, SaveFormat.Pdf);

                        Console.WriteLine($"Converted: {htmlPath} -> {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        // Log any errors for the specific file
                        Console.WriteLine($"Error converting {htmlPath}: {ex.Message}");
                    }
                });

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
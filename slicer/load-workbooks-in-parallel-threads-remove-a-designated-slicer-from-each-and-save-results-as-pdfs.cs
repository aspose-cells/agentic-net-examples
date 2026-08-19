// Title: Parallel removal of slicers from multiple Excel workbooks and PDF export using Aspose.Cells for .NET (C#)
// Description: A C# console app that loads a list of Excel files, runs each workbook on a separate thread, deletes the first slicer on every worksheet (if present), and saves the modified workbook as a PDF with Aspose.Cells PdfSaveOptions. Includes file‑existence checks, exception handling, and a final summary message.
// Keywords: Aspose.Cells | C# parallel Excel processing | remove slicer programmatically | Excel slicer collection | batch Excel to PDF conversion | multi‑threaded workbook conversion | PdfSaveOptions Aspose.Cells | delete slicers from worksheets | Aspose.Cells SlicerCollection | Excel to PDF C#
// Common Searches: how to delete slicers from Excel using Aspose.Cells C# | parallel processing of multiple workbooks with Aspose.Cells | batch convert Excel files to PDF in .NET | remove first slicer from each sheet before PDF export | Aspose.Cells multi‑threaded PDF conversion example
// Developer Intent: The developer needs to strip a specific slicer from every worksheet of several Excel workbooks concurrently and generate PDF versions of the cleaned files.
// Use Cases: Automated cleanup of reporting workbooks by removing slicers before publishing PDFs. | High‑throughput service that receives Excel uploads, removes slicer UI elements, and returns PDF archives. | CI/CD pipeline step that converts Excel templates to PDF after programmatically deleting slicers.
// AI Prompts: Write a reusable C# method that accepts a collection of Excel file paths, removes all slicers from each worksheet using Aspose.Cells, and returns the generated PDF file paths. | Show sample code for processing Excel workbooks in parallel, deleting the first slicer on each sheet, and saving each workbook as a PDF with custom PdfSaveOptions. | Explain best practices for exception handling, logging, and thread safety when removing slicers and converting workbooks to PDF with Aspose.Cells in a multi‑threaded application.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers; // Required for SlicerCollection

namespace AsposeCellsParallelSlicerRemoval
{
    // A C# console app that loads a list of Excel files, runs each workbook on a separate thread, deletes the first slicer on every worksheet (if present), and saves the modified workbook as a PDF with Aspose.Cells PdfSaveOptions. Includes file‑existence checks, exception handling, and a final summary message.
    class Program
    {
        static void Main()
        {
            // Input Excel files to process
            List<string> inputFiles = new List<string>
            {
                @"C:\Input\Workbook1.xlsx",
                @"C:\Input\Workbook2.xlsx",
                @"C:\Input\Workbook3.xlsx"
                // Add more file paths as needed
            };

            // Output directory for PDF files
            string outputDir = @"C:\Output\PDFs";
            Directory.CreateDirectory(outputDir);

            // Process each workbook in parallel
            List<Task> tasks = new List<Task>();
            foreach (string inputPath in inputFiles)
            {
                tasks.Add(Task.Run(() =>
                {
                    try
                    {
                        // Verify the input file exists
                        if (!File.Exists(inputPath))
                        {
                            Console.WriteLine($"Input file not found: {inputPath}");
                            return;
                        }

                        // Load the workbook from file
                        Workbook workbook = new Workbook(inputPath);

                        // Iterate through all worksheets and remove the first slicer if present
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            SlicerCollection slicers = sheet.Slicers;
                            if (slicers != null && slicers.Count > 0)
                            {
                                slicers.RemoveAt(0);
                            }
                        }

                        // Prepare PDF save options
                        PdfSaveOptions pdfOptions = new PdfSaveOptions();

                        // Determine output PDF file name
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                        string pdfPath = Path.Combine(outputDir, fileNameWithoutExt + ".pdf");

                        // Save the modified workbook as PDF
                        workbook.Save(pdfPath, pdfOptions);
                        Console.WriteLine($"Processed and saved PDF: {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{inputPath}': {ex.Message}");
                    }
                }));
            }

            // Wait for all parallel tasks to complete
            Task.WaitAll(tasks.ToArray());

            Console.WriteLine("All workbooks have been processed and saved as PDFs.");
        }
    }
}

// Title: Remove the first slicer from each worksheet in multiple Excel workbooks concurrently and export them as PDFs using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a collection of .xlsx files, uses Parallel.ForEach to delete the first slicer on every worksheet, and saves each workbook as a PDF with Aspose.Cells. | Update the parallel loop to remove all slicers from each worksheet before converting the workbook to PDF using Aspose.Cells. | Add comprehensive error handling and logging to the parallel slicer‑removal routine, recording which files succeeded or failed during PDF conversion.
// Common Searches: how to delete slicers from Excel worksheets in C# with Aspose.Cells before PDF export | parallel processing of multiple Excel files to PDF using Aspose.Cells .NET | remove first slicer from each sheet in a workbook programmatically Aspose.Cells | convert Excel workbooks to PDF after removing slicers with Aspose.Cells in a multi‑threaded application
// Tags: parallel slicer removal Aspose.Cells | excel worksheet slicer deletion C# | workbook to PDF conversion Aspose.Cells | multi‑threaded Excel processing .NET | Aspose.Cells PDFSaveOptions usage

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsParallelSlicerRemoval
{
    // // Loads a list of Excel files, removes the first slicer from every worksheet in parallel, and saves each workbook as a PDF using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel files (adjust paths as needed)
            List<string> excelFiles = new List<string>
            {
                @"Input\Workbook1.xlsx",
                @"Input\Workbook2.xlsx",
                @"Input\Workbook3.xlsx"
            };

            // Output directory for PDFs
            string outputDir = @"Output";
            Directory.CreateDirectory(outputDir);

            // Process each workbook in parallel
            Parallel.ForEach(excelFiles, excelPath =>
            {
                try
                {
                    // Verify the input file exists
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        return;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Remove the first slicer from each worksheet, if any
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        var slicers = sheet.Slicers;
                        if (slicers != null && slicers.Count > 0)
                        {
                            slicers.RemoveAt(0);
                        }
                    }

                    // Prepare PDF save options (default options are sufficient)
                    PdfSaveOptions pdfOptions = new PdfSaveOptions();

                    // Build output PDF file name based on input file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputDir, pdfFileName);

                    // Save the modified workbook as PDF
                    workbook.Save(pdfPath, pdfOptions);

                    Console.WriteLine($"Processed and saved PDF: {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            });

            Console.WriteLine("All files have been processed.");
        }
    }
}

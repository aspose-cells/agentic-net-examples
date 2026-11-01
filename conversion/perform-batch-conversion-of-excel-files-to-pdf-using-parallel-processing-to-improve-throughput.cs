// Title: Convert a folder of Excel workbooks (xls, xlsx, xlsm) to PDF in parallel with Aspose.Cells for .NET
// Common Searches: C# code to batch convert Excel files to PDF using Aspose.Cells and Parallel.ForEach | How to speed up Excel to PDF conversion with multithreading in .NET | Aspose.Cells example for converting multiple workbooks to PDF in parallel | Best practice for converting large numbers of Excel files to PDF in C# | Using MaxDegreeOfParallelism with Aspose.Cells workbook.Save to PDF
// Tags: parallel batch Excel to PDF conversion Aspose.Cells | Aspose.Cells Parallel.ForEach workbook export | C# multithreaded Excel workbook PDF generation | convert xls xlsx xlsm files to PDF .NET | maxdegreeofparallelism Aspose.Cells saveformat pdf

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace BatchExcelToPdf
{
    // The sample scans a directory for .xls, .xlsx, and .xlsm files, then uses Parallel.ForEach with a thread count equal to the processor cores to load each workbook via Aspose.Cells and save it as a PDF in the target folder, logging each conversion and handling errors without halting the batch process.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // var inputFiles = Directory.GetFiles(@"C:\ExcelFiles", "*.xlsx");
            // ConvertAllToPdf(inputFiles, @"C:\PdfOutput");

            // For demonstration, replace with actual paths.
            string inputFolder = @"C:\ExcelFiles";
            string outputFolder = @"C:\PdfOutput";

            // Ensure output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (xls, xlsx, xlsm) in the input folder.
            var inputFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            var excelFiles = new List<string>();
            foreach (var file in inputFiles)
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm")
                {
                    excelFiles.Add(file);
                }
            }

            // Perform parallel conversion.
            ConvertAllToPdf(excelFiles, outputFolder);
        }

        /// <param name="excelFilePaths">Full paths of the source Excel files.</param>
        /// <param name="outputDirectory">Directory where PDF files will be saved.</param>
        static void ConvertAllToPdf(IEnumerable<string> excelFilePaths, string outputDirectory)
        {
            // ParallelOptions can be tuned (e.g., MaxDegreeOfParallelism) if needed.
            ParallelOptions options = new ParallelOptions
            {
                // Use as many threads as there are logical processors.
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach(excelFilePaths, options, excelPath =>
            {
                try
                {
                    // Load the Excel workbook.
                    Workbook workbook = new Workbook(excelPath); // Load rule

                    // Determine output PDF file name.
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputDirectory, pdfFileName);

                    // Save the workbook as PDF.
                    workbook.Save(pdfPath, SaveFormat.Pdf); // Save rule

                    Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    // Log the error but continue processing other files.
                    Console.Error.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            });
        }
    }
}

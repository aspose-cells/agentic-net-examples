// Title: Convert a folder of Excel workbooks (.xls, .xlsx, .xlsm) to PDF in parallel with Aspose.Cells and C#
// AI Prompts: Write a C# console program that scans an input directory recursively for .xls, .xlsx, and .xlsm files and converts each workbook to PDF using Aspose.Cells inside a Parallel.ForEach loop. | Demonstrate how to keep the original folder hierarchy when writing the resulting PDFs to a separate output directory. | Implement per‑file exception handling that logs conversion errors without stopping the parallel batch operation.
// Common Searches: how to use Aspose.Cells to convert multiple Excel files to PDF in C# with parallel processing | preserve source folder structure when exporting Excel workbooks to PDF using .NET | batch convert .xls .xlsx .xlsm to PDF with multithreading in a console app | example of Parallel.ForEach for Excel to PDF conversion with Aspose.Cells | C# program to recursively find Excel files and save them as PDFs
// Tags: parallel Aspose.Cells Excel to PDF conversion | recursive Excel file enumeration C# | preserve directory hierarchy PDF output | PdfSaveOptions usage Aspose.Cells | multi-threaded workbook conversion .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ExcelToPdfBatch
{
    // A C# console utility that accepts input and output folder paths, recursively discovers .xls, .xlsx, and .xlsm files, and converts each workbook to PDF using Aspose.Cells. The conversion runs inside Parallel.ForEach with a degree of parallelism matching the processor count, preserving the original directory structure in the output location and logging any file‑specific errors without halting the batch process.
    class Program
    {
        static void Main(string[] args)
        {
            // args[0] = input folder path, args[1] = output folder path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExcelToPdfBatch <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Find all Excel files recursively
            var excelFiles = Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                .ToList();

            Console.WriteLine($"Found {excelFiles.Count} Excel file(s) to convert.");

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach(excelFiles, parallelOptions, excelPath =>
            {
                try
                {
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        return;
                    }

                    // Load workbook
                    var workbook = new Workbook(excelPath);

                    // PDF save options (customize if needed)
                    var pdfOptions = new PdfSaveOptions();

                    // Preserve relative folder structure in output
                    string relativePath = Path.GetRelativePath(inputFolder, excelPath);
                    string pdfRelativePath = Path.ChangeExtension(relativePath, ".pdf");
                    string pdfFullPath = Path.Combine(outputFolder, pdfRelativePath);

                    // Ensure target directory exists
                    string? pdfDirectory = Path.GetDirectoryName(pdfFullPath);
                    if (!string.IsNullOrEmpty(pdfDirectory) && !Directory.Exists(pdfDirectory))
                    {
                        Directory.CreateDirectory(pdfDirectory);
                    }

                    // Save as PDF
                    workbook.Save(pdfFullPath, pdfOptions);
                    Console.WriteLine($"Converted: {excelPath} -> {pdfFullPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            });

            Console.WriteLine("Batch conversion completed.");
        }
    }
}

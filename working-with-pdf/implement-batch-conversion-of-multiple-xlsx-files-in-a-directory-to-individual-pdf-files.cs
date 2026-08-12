// Title: Batch convert multiple XLSX files to PDF with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a given folder for *.xlsx workbooks, creates an output directory if needed, loads each workbook with Aspose.Cells, and saves it as a PDF with the same base name. Includes defensive checks, error handling, and progress logging for reliable batch conversion.
// Keywords: Aspose.Cells batch conversion | XLSX to PDF C# | convert multiple Excel files to PDF | Aspose.Cells save workbook as PDF | directory based Excel PDF conversion | C# console utility Excel PDF | bulk Excel to PDF .NET
// Common Searches: batch convert xlsx to pdf asp.net | c# code to convert all excel files in a folder to pdf | aspocells convert multiple workbooks to pdf | how to automate excel to pdf conversion with aspocells | command line tool for bulk xlsx pdf conversion
// Developer Intent: Automatically transform every .xlsx file in a specified directory into an individual PDF using Aspose.Cells.
// Use Cases: Nightly job that turns a folder of Excel reports into PDFs for archiving. | Command‑line tool for processing user‑uploaded Excel files and delivering PDFs to downstream systems. | Web API endpoint that accepts a zip of XLSX files, runs batch conversion, and returns a zip of PDFs.
// AI Prompts: Generate C# code that iterates over all .xlsx files in a directory and uses Aspose.Cells to save each as a PDF with matching filenames. | Explain best practices for error handling and logging when batch converting Excel workbooks to PDF with Aspose.Cells. | Create a PowerShell script that calls a compiled .NET executable to perform bulk XLSX‑to‑PDF conversion using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // A C# console utility that scans a given folder for *.xlsx workbooks, creates an output directory if needed, loads each workbook with Aspose.Cells, and saves it as a PDF with the same base name. Includes defensive checks, error handling, and progress logging for reliable batch conversion.
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing XLSX files
            string inputDir = @"C:\InputXlsx";
            // Output directory where PDF files will be saved
            string outputDir = @"C:\OutputPdf";

            // Verify that the input directory exists
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Input directory not found: {inputDir}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            try
            {
                // Get all .xlsx files in the input directory (non‑recursive)
                string[] xlsxFiles = Directory.GetFiles(inputDir, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string xlsxPath in xlsxFiles)
                {
                    // Verify the source file exists (defensive check)
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    try
                    {
                        // Build the PDF file name based on the source file name
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                        string pdfPath = Path.Combine(outputDir, fileNameWithoutExt + ".pdf");

                        // Load the workbook and save it as PDF
                        Workbook workbook = new Workbook(xlsxPath);
                        workbook.Save(pdfPath, SaveFormat.Pdf);

                        Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(pdfPath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}

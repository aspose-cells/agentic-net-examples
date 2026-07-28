// Title: C# Console App to Batch Convert Excel to PDF using Aspose.Cells
// Description: A command‑line utility that receives a folder path, scans the directory for Excel workbooks, and converts each file to a PDF with the same name. Non‑Excel files and existing PDFs are ignored, and conversion results are logged to the console.
// Keywords: Aspose.Cells batch PDF conversion | C# console Excel to PDF | folder based workbook conversion | ConversionUtility.Convert example | command line Excel PDF tool | process multiple Excel files C# | automated Excel PDF generation
// Common Searches: how to convert all Excel files in a folder to PDF with Aspose.Cells | C# batch Excel to PDF command line utility | Aspose.Cells convert workbook to PDF without opening Excel | script to export multiple spreadsheets as PDFs | automate Excel PDF conversion using Aspose
// Developer Intent: Create a lightweight C# console program that enumerates a given directory, identifies Excel workbooks via Aspose.Cells.FileFormatUtil, and uses Aspose.Cells.Utility.ConversionUtility to generate matching PDF files while handling errors gracefully.
// Use Cases: Nightly generation of PDF reports from a repository of Excel spreadsheets. | Providing end‑users a no‑Office tool to convert many workbooks to PDF in one click. | Embedding the conversion step in CI/CD pipelines to validate visual output of Excel files. | Archiving Excel data as PDFs for compliance without manual intervention.
// AI Prompts: Generate a C# console application that accepts a folder path argument and batch converts every Excel workbook in that folder to PDF using Aspose.Cells, skipping non‑Excel files. | Show how to log each successful conversion and capture exceptions when processing multiple workbooks with Aspose.Cells. | Provide sample code that uses Aspose.Cells.FileFormatUtil to detect supported Excel formats and Aspose.Cells.Utility.ConversionUtility to create PDF files with identical base names.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ExcelToPdfBatch
{
    // A command‑line utility that receives a folder path, scans the directory for Excel workbooks, and converts each file to a PDF with the same name. Non‑Excel files and existing PDFs are ignored, and conversion results are logged to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect a folder path as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToPdfBatch <folderPath>");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: The folder \"{folderPath}\" does not exist.");
                return;
            }

            // Process each file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                try
                {
                    // Determine if the file is an Excel workbook based on its extension
                    string ext = Path.GetExtension(filePath);
                    if (string.IsNullOrEmpty(ext))
                        continue;

                    // Remove the leading dot and convert to lower case
                    string extWithoutDot = ext.TrimStart('.').ToLowerInvariant();

                    // Map the extension to a SaveFormat; if not recognized, skip the file
                    SaveFormat format = FileFormatUtil.ExtensionToSaveFormat(extWithoutDot);
                    if (format == SaveFormat.Unknown || format == SaveFormat.Pdf)
                        continue; // Not an Excel format or already a PDF

                    // Build the output PDF file name
                    string pdfPath = Path.ChangeExtension(filePath, ".pdf");

                    // Ensure the source file still exists before conversion
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: \"{filePath}\"");
                        continue;
                    }

                    // Convert the workbook to PDF
                    ConversionUtility.Convert(filePath, pdfPath);

                    Console.WriteLine($"Converted: \"{Path.GetFileName(filePath)}\" → \"{Path.GetFileName(pdfPath)}\"");
                }
                catch (Exception ex)
                {
                    // Log conversion errors but continue processing other files
                    Console.WriteLine($"Failed to convert \"{Path.GetFileName(filePath)}\": {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}

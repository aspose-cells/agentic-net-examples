// Title: Convert all XLSX files in a folder to separate PDF documents with Aspose.Cells for .NET (C# batch example)
// AI Prompts: Write a C# console program that enumerates every .xlsx file in a specified directory, loads each workbook with Aspose.Cells, and saves it as a PDF using SaveFormat.Pdf. | Enhance the batch converter to catch missing files, load failures, or save errors, log the details, and continue processing the remaining workbooks. | Update the script so it automatically creates the output folder when it does not already exist before writing PDF files.
// Common Searches: c# Aspose.Cells batch convert multiple xlsx files to pdf in a single folder | how to loop through a directory of Excel workbooks and export each to PDF using Aspose.Cells | sample code for converting all .xlsx files in a folder to PDF with Aspose.Cells .NET | Aspose.Cells SaveFormat.Pdf batch processing example c#
// Tags: Aspose.Cells batch conversion XLSX → PDF | C# file system enumeration for Excel to PDF | Workbook.Save with SaveFormat.Pdf in loop | auto-create output directory C# | robust error handling for batch workbook conversion

using System;
using System.IO;
using Aspose.Cells;

// The example scans a given input folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an individual PDF in an output folder. It creates the output directory if missing and includes error handling to log and skip problematic files while continuing the batch operation.
class BatchXlsxToPdfConverter
{
    static void Main(string[] args)
    {
        // Define the directories for input XLSX files and output PDFs.
        string inputDirectory = @"C:\InputXlsx";
        string outputDirectory = @"C:\OutputPdf";

        // Verify that the input directory exists; if not, inform the user and exit.
        if (!Directory.Exists(inputDirectory))
        {
            Console.WriteLine($"Input directory not found: {inputDirectory}");
            return;
        }

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        string[] xlsxFiles;
        try
        {
            // Retrieve all .xlsx files in the input directory (non‑recursive).
            xlsxFiles = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input directory: {ex.Message}");
            return;
        }

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Verify the file still exists before loading.
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found (skipped): {Path.GetFileName(xlsxPath)}");
                    continue;
                }

                // Load the workbook from the XLSX file.
                Workbook workbook = new Workbook(xlsxPath);

                // Determine the output PDF file name.
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string pdfPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".pdf");

                // Save the workbook as PDF.
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(pdfPath)}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files.
                Console.WriteLine($"Error converting '{Path.GetFileName(xlsxPath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}

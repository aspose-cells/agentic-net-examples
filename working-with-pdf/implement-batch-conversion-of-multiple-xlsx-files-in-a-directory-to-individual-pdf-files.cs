// Title: Convert every XLSX workbook in a folder to a separate PDF using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as a PDF with the same base name in an output folder. | Add per‑file try/catch logic so that a conversion failure logs the file name and error message while the batch continues processing the remaining workbooks. | Extend the script to recreate the source directory’s sub‑folder structure inside the destination folder when writing the PDF files.
// Common Searches: aspocells c# batch convert excel files in a directory to pdf | how to loop through all .xlsx files and export each to pdf using Aspose.Cells | c# program to convert multiple Excel workbooks to PDF and keep original filenames
// Tags: Aspose.Cells multiple workbook PDF export | C# enumerate .xlsx files in folder | Workbook.Save with SaveFormat.Pdf | individual file error logging in conversion loop | maintain source directory hierarchy for PDF output

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that enumerates every .xlsx file in a given input folder, loads each workbook with Aspose.Cells, and saves it as a PDF with the same filename in a designated output folder. The program creates the output directory if needed, logs conversion errors per file, and can be extended to preserve the original sub‑directory layout.
class BatchXlsxToPdfConverter
{
    static void Main()
    {
        // Define the source directory containing XLSX files
        string sourceDirectory = @"C:\InputXlsxFiles";

        // Define the target directory where PDF files will be saved
        string targetDirectory = @"C:\OutputPdfFiles";

        // Verify source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory not found: {sourceDirectory}");
            return;
        }

        // Ensure the target directory exists
        if (!Directory.Exists(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }

        try
        {
            // Get all XLSX files in the source directory
            string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Load the XLSX workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Build the PDF file path (same name, .pdf extension)
                    string pdfFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                    string pdfPath = Path.Combine(targetDirectory, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    // Log errors for individual files without stopping the batch
                    Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors that may occur during enumeration
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

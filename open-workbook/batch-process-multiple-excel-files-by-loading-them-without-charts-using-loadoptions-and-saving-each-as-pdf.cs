// Title: C# batch conversion of Excel (.xlsx/.xls) workbooks to PDF using Aspose.Cells
// AI Prompts: Create a C# console application that scans a specified directory, loads each .xlsx or .xls workbook with Aspose.Cells, and saves it as a PDF in an output folder, including error handling for missing folders. | Write C# code that enumerates all Excel files in a folder, skips non‑Excel files, ensures the destination directory exists, and uses Workbook.Save to export each workbook to PDF in a loop.
// Common Searches: c# aspocells batch convert excel files in a folder to pdf | how to iterate over multiple xlsx files and export to pdf using aspocells | aspocells save workbook as pdf for each file in directory c# | console app bulk excel to pdf conversion aspocells .net
// Tags: aspocells batch workbook to pdf conversion | c# load excel files and export to pdf with aspocells | process multiple xlsx files aspocells saveformat pdf | aspocells workbook.save pdf in loop | c# console bulk file conversion aspocells

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that iterates through a given input folder, loads each .xlsx or .xls workbook with Aspose.Cells, and saves it as a PDF in an output directory, creating the output folder if necessary and skipping non‑Excel files.
class BatchExcelToPdf
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string inputFolder = @"C:\InputExcelFiles";
        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\OutputPdfFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx or .xls file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls")
                continue; // Skip non‑Excel files

            try
            {
                // Load the workbook (default loading includes all objects)
                Workbook workbook = new Workbook(filePath);

                // Build the output PDF file name
                string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {pdfFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}

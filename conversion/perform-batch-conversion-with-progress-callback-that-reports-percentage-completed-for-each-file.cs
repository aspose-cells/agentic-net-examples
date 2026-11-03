// Title: Batch convert multiple .xlsx workbooks to PDF with Aspose.Cells for .NET (C#)
// Common Searches: c# console app batch convert all .xlsx files in a directory to PDF using Aspose.Cells | how to use Aspose.Cells Workbook.Save to generate PDFs from multiple Excel workbooks | example code for iterating over Excel files and saving each as PDF with Aspose.Cells | Aspose.Cells PDFSaveOptions usage in a folder‑wide Excel to PDF conversion | C# error handling pattern for bulk Excel to PDF conversion with Aspose.Cells
// Tags: Aspose.Cells batch conversion to PDF | C# iterate over .xlsx files | Workbook.Save with PdfSaveOptions | Console application Excel to PDF processing | Aspose.Cells conversion error handling | Folder based workbook conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// A C# console program that scans a specified folder for *.xlsx files, creates an output directory if needed, and converts each workbook to PDF using Aspose.Cells Workbook and PdfSaveOptions, with basic file existence checks and exception handling.
class BatchConverter
{
    // Converts a single workbook to PDF with basic error handling
    private static void ConvertToPdf(string sourcePath, string destPath)
    {
        try
        {
            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the source file
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options (no progress handler to maintain compatibility)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(destPath, saveOptions);
            Console.WriteLine($"Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
        }
    }

    static void Main(string[] args)
    {
        // Folder containing source Excel files
        string inputFolder = @"C:\InputFiles";
        // Folder where converted PDFs will be placed
        string outputFolder = @"C:\OutputFiles";

        try
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all Excel files (adjust the pattern as needed)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string excelFile in excelFiles)
            {
                // Build the output PDF file path
                string pdfFileName = Path.GetFileNameWithoutExtension(excelFile) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Perform conversion
                ConvertToPdf(excelFile, pdfPath);
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

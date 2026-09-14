// Title: Batch convert a directory of .xls and .xlsx files to single-page-per-sheet PDFs with Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that scans a specified folder for .xls and .xlsx files, loads each workbook with Aspose.Cells, and saves it as a PDF using PdfSaveOptions with OnePagePerSheet enabled. | Show how to configure PdfSaveOptions in Aspose.Cells to enforce one PDF page per worksheet when exporting multiple Excel workbooks in a batch process. | Write code that creates input and output directories, iterates over Excel files, converts each to PDF, and logs success or error messages for every file.
// Common Searches: asp.net batch convert excel files to pdf one page per sheet aspose.cells | c# script to convert all .xlsx files in a folder to pdf using aspose cells | how to set OnePagePerSheet option when exporting Excel to PDF in .NET | automate conversion of multiple workbooks to pdf with aspose.cells console app
// Tags: batch excel to pdf conversion aspose.cells | onepagepersheet pdfsaveoptions .net | c# folder processing aspose.cells | excel workbook pdf export aspose.cells | automated pdf generation from excel files

using System;
using System.IO;
using Aspose.Cells;

// The C# console application scans an input directory for .xls and .xlsx workbooks, creates PdfSaveOptions with OnePagePerSheet set to true, and saves each workbook as a PDF in an output folder. It ensures the output folders exist, skips non‑Excel files, logs conversion results, and continues processing even if individual files cause errors.
class ExcelToPdfConverter
{
    static void Main(string[] args)
    {
        try
        {
            // Input folder containing Excel workbooks
            string inputFolder = @"C:\InputExcelFiles";
            // Output folder for generated PDFs
            string outputFolder = @"C:\OutputPdfFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                Console.WriteLine("Please create the folder and add Excel files before running the program.");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Define PDF save options with OnePagePerSheet enabled
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Process each Excel file in the input folder (supports .xls and .xlsx)
            foreach (string excelFilePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string extension = Path.GetExtension(excelFilePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non-Excel files

                // Verify the file exists before attempting to load
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine($"File not found: {excelFilePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelFilePath);

                    // Determine the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelFilePath) + ".pdf";
                    string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF with the specified options
                    workbook.Save(pdfFilePath, pdfOptions);
                    Console.WriteLine($"Converted: {excelFilePath} -> {pdfFilePath}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error processing '{excelFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

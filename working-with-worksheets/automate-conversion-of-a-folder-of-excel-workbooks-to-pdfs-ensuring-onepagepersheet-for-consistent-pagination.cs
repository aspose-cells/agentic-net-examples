// Title: C# batch conversion of Excel workbooks to PDF with OnePagePerSheet using Aspose.Cells
// Description: A complete C# example that scans a source folder for Excel files (*.xls, *.xlsx, *.xlsm, *.xlsb), loads each workbook with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet = true, and saves the result as a PDF in an output folder. The code creates missing directories, handles errors gracefully, and can be run from the command line or integrated into automated workflows.
// Keywords: Aspose.Cells | C# batch Excel to PDF | OnePagePerSheet | PdfSaveOptions | convert folder of Excel files | Excel to PDF automation | process multiple workbooks | .NET PDF pagination | Aspose.Cells example | GitHub sample code
// Common Searches: Aspose.Cells convert all Excel files in a folder to PDF C# | OnePagePerSheet option Aspose.Cells example | Batch Excel to PDF conversion .NET | C# code to process multiple workbooks with Aspose.Cells | Create PDF from Excel folder using Aspose.Cells | Automate Excel to PDF conversion with pagination
// Developer Intent: Automatically convert every Excel workbook in a specified directory to a PDF file, ensuring each worksheet is rendered on a single page.
// Use Cases: Generate printable PDFs for a collection of monthly financial reports stored in a shared drive. | Archive project spreadsheets as paginated PDFs before uploading to a document‑management system. | Run nightly jobs that transform incoming Excel uploads into PDFs for downstream processing or compliance.
// AI Prompts: Write a C# method that recursively scans a directory and its subfolders, converting each Excel file to PDF with Aspose.Cells while logging successes and failures to a CSV file. | Explain how to add password protection to the PDFs produced by the batch conversion code without affecting the OnePagePerSheet layout. | Provide a PowerShell script that invokes the compiled C# batch converter, passing source and destination paths as arguments.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchConversion
{
    // A complete C# example that scans a source folder for Excel files (*.xls, *.xlsx, *.xlsm, *.xlsb), loads each workbook with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet = true, and saves the result as a PDF in an output folder. The code creates missing directories, handles errors gracefully, and can be run from the command line or integrated into automated workflows.
    public static class ExcelToPdfConverter
    {
        /// <param name="sourceFolder">Folder containing Excel files.</param>
        /// <param name="outputFolder">Folder where PDF files will be written.</param>
        public static void Run(string sourceFolder, string outputFolder)
        {
            // Ensure the source directory exists
            if (!Directory.Exists(sourceFolder))
            {
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");
            }

            // Create the output directory if it does not exist
            Directory.CreateDirectory(outputFolder);

            // Search for common Excel extensions
            string[] excelPatterns = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };

            foreach (string pattern in excelPatterns)
            {
                foreach (string excelFilePath in Directory.GetFiles(sourceFolder, pattern))
                {
                    // Verify the file still exists before processing
                    if (!File.Exists(excelFilePath))
                    {
                        Console.WriteLine($"File not found (skipped): {excelFilePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(excelFilePath);

                        // Configure PDF save options
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            OnePagePerSheet = true
                        };

                        // Build the output PDF file path
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelFilePath) + ".pdf";
                        string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

                        // Save the workbook as PDF
                        workbook.Save(pdfFilePath, pdfOptions);

                        Console.WriteLine($"Converted: {excelFilePath} -> {pdfFilePath}");
                    }
                    catch (Exception ex)
                    {
                        // Log conversion errors but continue processing other files
                        Console.WriteLine($"Error converting '{excelFilePath}': {ex.Message}");
                    }
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string sourceFolder;
            string outputFolder;

            if (args.Length >= 2)
            {
                sourceFolder = args[0];
                outputFolder = args[1];
            }
            else
            {
                // Default folders relative to the executable location
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                sourceFolder = Path.Combine(baseDir, "Input");
                outputFolder = Path.Combine(baseDir, "Output");
            }

            try
            {
                ExcelToPdfConverter.Run(sourceFolder, outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}

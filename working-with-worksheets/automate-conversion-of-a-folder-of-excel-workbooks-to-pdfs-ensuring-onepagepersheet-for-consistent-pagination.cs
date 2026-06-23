using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel workbooks
            string sourceFolder = @"C:\InputExcels";

            // Folder where the resulting PDF files will be saved
            string outputFolder = @"C:\OutputPDFs";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            try
            {
                // Get all Excel files in the source folder (supports common extensions)
                string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string excelPath in excelFiles)
                {
                    // Filter only Excel workbook files based on extension
                    string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                        continue;

                    // Verify the file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found (skipped): {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file
                        Workbook workbook = new Workbook(excelPath);

                        // Configure PDF save options to force one page per sheet
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            OnePagePerSheet = true
                        };

                        // Determine the output PDF file path
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Save the workbook as PDF using the specified options
                        workbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"Converted '{excelPath}' to '{pdfPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
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
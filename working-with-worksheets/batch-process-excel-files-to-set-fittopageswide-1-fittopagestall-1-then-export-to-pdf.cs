using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchFitToPagesAndExportPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputExcels";
            // Folder where PDF files will be saved
            string outputFolder = @"C:\OutputPdfs";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Process each Excel file in the source folder
                foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
                {
                    try
                    {
                        // Verify the Excel file exists before loading
                        if (!File.Exists(excelPath))
                        {
                            Console.WriteLine($"File not found (skipped): {excelPath}");
                            continue;
                        }

                        // Load the workbook
                        Workbook workbook = new Workbook(excelPath);

                        // Set FitToPagesWide = 1 and FitToPagesTall = 1 for every worksheet
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            sheet.PageSetup.FitToPagesWide = 1;
                            sheet.PageSetup.FitToPagesTall = 1;
                        }

                        // Prepare PDF save options (default options are sufficient)
                        PdfSaveOptions pdfOptions = new PdfSaveOptions();

                        // Determine output PDF file name
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Save the workbook as PDF
                        workbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"Converted '{excelPath}' to PDF at '{pdfPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
// Title: Batch convert multiple Excel files to PDF with FitToPagesWide=1 and FitToPagesTall=1 using Aspose.Cells for .NET
// Common Searches: how to convert all .xlsx files in a folder to PDF with fit to one page using Aspose.Cells C# | set page setup FitToPagesWide and FitToPagesTall for every worksheet before saving as PDF in Aspose.Cells | skip password‑protected Excel workbooks while batch exporting to PDF with Aspose.Cells .NET | batch process Excel files to PDF with Aspose.Cells and preserve one‑page layout | Aspose.Cells C# export workbook to PDF with page scaling set to 1 page wide and tall
// Tags: batch Excel to PDF conversion Aspose.Cells | set FitToPagesWide FitToPagesTall Aspose.Cells | page setup scaling for PDF export C# | skip password protected workbooks Aspose.Cells | process folder of .xlsx files programmatically

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified directory for .xlsx files, loads each workbook with Aspose.Cells, applies a page setup that fits each worksheet to one page wide and tall, saves the workbook as a PDF with the same name, skips password‑protected files, and logs any processing errors.
class BatchExcelToPdf
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string sourceFolder = @"C:\ExcelFiles";

        // Ensure the folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Get all Excel files (you can add other extensions if needed)
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string excelPath in excelFiles)
        {
            // Verify the file still exists before attempting to load
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the workbook (no password)
                Workbook workbook = new Workbook(excelPath);

                // Set each worksheet to fit to one page wide and tall
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    PageSetup pageSetup = sheet.PageSetup;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = 1;
                }

                // Determine PDF output path (same name, .pdf extension)
                string pdfPath = Path.ChangeExtension(excelPath, ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Converted '{Path.GetFileName(excelPath)}' to PDF.");
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                // Skip password‑protected files
                Console.WriteLine($"Skipping password‑protected file: {Path.GetFileName(excelPath)}");
            }
            catch (Exception ex)
            {
                // Log any other errors and continue processing remaining files
                Console.WriteLine($"Error processing '{Path.GetFileName(excelPath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}

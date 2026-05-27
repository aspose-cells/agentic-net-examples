using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public class WorkbookBatchProcessor
{
    // Processes all Excel files in the specified input folder.
    // For each workbook:
    //   - Hides zero values (DisplayZeros = false)
    //   - Sets worksheet zoom to 80%
    //   - Saves the workbook as PDF in the output folder.
    public static void ProcessWorkbooks(string inputFolder, string outputFolder)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (XLS, XLSX) in the input folder.
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files.

            // Load the workbook (create rule).
            Workbook workbook = new Workbook(filePath);

            // Apply settings to each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Hide zero values.
                sheet.DisplayZeros = false; // Worksheet.DisplayZeros property.

                // Apply 80 % zoom.
                sheet.PageSetup.Zoom = 80; // Worksheet.PageSetup.Zoom property.
            }

            // Prepare PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export only visible sheets (optional, uses PaginatedSaveOptions.SheetSet).
                SheetSet = SheetSet.Visible
            };

            // Build output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the workbook as PDF (save rule).
            workbook.Save(pdfPath, pdfOptions);
        }
    }

    // Example usage.
    public static void Main()
    {
        string inputFolder = @"C:\InputWorkbooks";
        string outputFolder = @"C:\OutputPdfs";

        ProcessWorkbooks(inputFolder, outputFolder);

        Console.WriteLine("Batch processing completed.");
    }
}
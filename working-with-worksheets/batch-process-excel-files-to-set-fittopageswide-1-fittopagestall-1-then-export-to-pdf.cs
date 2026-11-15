// Title: Batch convert multiple Excel workbooks to PDF with 1‑page‑wide and 1‑page‑tall scaling using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch convert .xls and .xlsx files to PDF with FitToPagesWide=1 and FitToPagesTall=1 | How to set page scaling for all worksheets before exporting Excel to PDF in Aspose.Cells | Programmatically process a folder of Excel files and save each as a PDF with single‑page layout using Aspose.Cells | Aspose.Cells example for applying FitToPages settings to every worksheet in a workbook | C# loop through Excel files and generate PDF with page fit settings using Aspose.Cells
// Tags: batch Excel to PDF conversion Aspose.Cells | set FitToPagesWide FitToPagesTall Aspose.Cells | apply page scaling to all worksheets C# | export workbook as PDF with single page layout | process multiple Excel files programmatically

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input folder for .xls and .xlsx files, loads each workbook with Aspose.Cells, sets each worksheet's PageSetup FitToPagesWide and FitToPagesTall properties to 1 (forcing a single‑page‑wide and single‑page‑tall layout), and saves the result as a PDF in an output folder.
class ExcelToPdfBatchProcessor
{
    static void Main(string[] args)
    {
        // Input folder containing Excel files
        string inputFolder = @"C:\InputExcelFiles";
        // Output folder for generated PDFs
        string outputFolder = @"C:\OutputPdfFiles";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (both .xls and .xlsx) in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Apply FitToPages settings to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
                pageSetup.FitToPagesTall = 1;   // Fit to 1 page tall
            }

            // Determine output PDF file name
            string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }

        Console.WriteLine("Batch processing completed.");
    }
}

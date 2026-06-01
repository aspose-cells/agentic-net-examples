using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering; // for PdfSaveOptions if needed

class Program
{
    static void Main()
    {
        // List of Excel files to process
        List<string> excelFiles = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx",
            "Book3.xlsx"
        };

        // Temporary folder for generated PDFs
        string pdfFolder = Path.Combine(Path.GetTempPath(), "AsposePdfOutput");
        Directory.CreateDirectory(pdfFolder);

        // Process each workbook
        foreach (string excelPath in excelFiles)
        {
            // Load workbook (uses Workbook(string) constructor rule)
            Workbook wb = new Workbook(excelPath);

            // Remove all slicers from every worksheet
            foreach (Worksheet ws in wb.Worksheets)
            {
                SlicerCollection slicers = ws.Slicers;
                // Remove slicers while collection is not empty
                while (slicers.Count > 0)
                {
                    // Delete slicer at index 0 (uses SlicerCollection.RemoveAt rule)
                    slicers.RemoveAt(0);
                }
            }

            // Prepare PDF file name
            string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
            string pdfPath = Path.Combine(pdfFolder, pdfFileName);

            // Save workbook as PDF (uses Workbook.Save(string, SaveFormat) rule)
            wb.Save(pdfPath, SaveFormat.Pdf);

            // Release resources
            wb.Dispose();
        }

        // Create a zip archive containing all PDFs
        string zipPath = "AllWorkbooks.pdf.zip";
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string pdfFile in Directory.GetFiles(pdfFolder, "*.pdf"))
            {
                // Add each PDF to the zip (free‑form code, no specific rule)
                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
            }
        }

        // Clean up temporary PDF files
        Directory.Delete(pdfFolder, true);

        Console.WriteLine("All workbooks have been converted to PDF, slicers removed, and archived to: " + zipPath);
    }
}
// Title: Remove all slicers from multiple Excel workbooks, convert each to PDF, and zip the PDFs using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a folder for *.xlsx files, removes every slicer from each worksheet with Aspose.Cells, saves each workbook as a PDF, and then compresses all PDFs into a single zip archive. | Generate .NET code that batch‑processes Excel workbooks: clear slicer collections, export the workbooks to PDF, and archive the resulting PDFs using System.IO.Compression.
// Common Searches: Aspose.Cells C# remove slicers from all worksheets in a workbook | batch convert Excel files to PDF and create a zip archive with .NET | clear slicer collections programmatically using Aspose.Cells | save workbook as PDF after deleting slicers Aspose.Cells | compress multiple PDF files into a zip file in C# after Excel conversion
// Tags: remove slicers Aspose.Cells C# | export workbook to PDF Aspose.Cells | batch Excel to PDF conversion .NET | zip multiple PDFs System.IO.Compression | clear worksheet slicer collections programmatically | temporary folder PDF generation Aspose

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The program iterates over all .xlsx files in a specified directory, loads each workbook with Aspose.Cells, clears every slicer from every worksheet, saves the modified workbook as a PDF in a temporary location, then packages all PDFs into a single ZIP archive and cleans up the temporary files.
class Program
{
    static void Main()
    {
        // Folder that contains the source Excel workbooks
        string sourceFolder = @"C:\InputWorkbooks";

        // Temporary folder to store intermediate PDF files
        string tempPdfFolder = Path.Combine(Path.GetTempPath(), "AsposePdfTemp");
        Directory.CreateDirectory(tempPdfFolder);

        // Collect all Excel files (adjust the pattern if needed)
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        // List to keep track of generated PDF file paths
        List<string> pdfFiles = new List<string>();

        foreach (string excelPath in excelFiles)
        {
            // Load the workbook from file (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(excelPath);

            // Remove all slicers from every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear removes all slicers in the collection
                sheet.Slicers.Clear();
            }

            // Save the modified workbook as PDF (uses the provided Save(string, SaveFormat) method)
            string pdfPath = Path.Combine(
                tempPdfFolder,
                Path.GetFileNameWithoutExtension(excelPath) + ".pdf");

            workbook.Save(pdfPath, SaveFormat.Pdf);
            pdfFiles.Add(pdfPath);

            // Release resources
            workbook.Dispose();
        }

        // Create a ZIP archive containing all generated PDFs
        string zipPath = Path.Combine(sourceFolder, "WorkbooksPdfArchive.zip");
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (string pdfFile in pdfFiles)
            {
                // Add each PDF to the archive with its file name
                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
            }
        }

        // Optional: clean up temporary PDF files
        foreach (string pdfFile in pdfFiles)
        {
            File.Delete(pdfFile);
        }
        Directory.Delete(tempPdfFolder, true);

        Console.WriteLine("All workbooks processed, slicers removed, PDFs archived to:");
        Console.WriteLine(zipPath);
    }
}

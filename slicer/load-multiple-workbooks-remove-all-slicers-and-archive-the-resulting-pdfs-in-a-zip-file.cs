using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Paths of the source Excel workbooks
        string[] sourceFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };

        // List to keep generated PDF file paths
        List<string> pdfFiles = new List<string>();

        // Process each workbook
        foreach (string srcPath in sourceFiles)
        {
            // Load the workbook (uses the provided Workbook(string) constructor)
            Workbook wb = new Workbook(srcPath);

            // Remove all slicers from every worksheet
            foreach (Worksheet ws in wb.Worksheets)
            {
                SlicerCollection slicers = ws.Slicers;
                // Remove slicers in reverse order to avoid index shifting
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    slicers.RemoveAt(i); // uses the provided RemoveAt method
                }
            }

            // Create a temporary PDF file name
            string tempPdf = Path.Combine(Path.GetTempPath(),
                                          Guid.NewGuid().ToString() + ".pdf");

            // Save the workbook as PDF (uses the provided Save(string, SaveFormat) method)
            wb.Save(tempPdf, SaveFormat.Pdf);

            // Keep track of the PDF for later archiving
            pdfFiles.Add(tempPdf);

            // Release resources
            wb.Dispose();
        }

        // Archive all PDFs into a single ZIP file
        string zipPath = "WorkbooksArchive.zip";
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (string pdfPath in pdfFiles)
            {
                // Add each PDF to the archive with its file name
                archive.CreateEntryFromFile(pdfPath, Path.GetFileName(pdfPath));
            }
        }

        // Clean up temporary PDF files
        foreach (string pdfPath in pdfFiles)
        {
            if (File.Exists(pdfPath))
            {
                File.Delete(pdfPath);
            }
        }

        Console.WriteLine($"All PDFs have been archived to '{zipPath}'.");
    }
}
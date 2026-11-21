// Title: Batch remove slicers from Excel workbooks and create a PDF ZIP archive using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove slicers from all worksheets in multiple Excel files | How to delete slicers and convert Excel workbooks to PDF with Aspose.Cells | Batch process Excel files to strip slicers and compress PDFs into a zip in .NET | Aspose.Cells generate PDFs from workbooks after removing slicers and archive them | C# script to remove slicers, save as PDF, and zip PDFs using System.IO.Compression
// Tags: remove slicers Aspose.Cells C# | export workbook to PDF Aspose.Cells | zip multiple PDFs .NET | batch process Excel workbooks Aspose.Cells | temporary PDF folder C#

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The C# program scans a folder for .xlsx files, loads each workbook with Aspose.Cells, removes every slicer from all worksheets, saves the cleaned workbook as a PDF in a temporary directory, then packages all PDFs into a ZIP archive using System.IO.Compression and finally deletes the temporary folder.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string sourceFolder = @"C:\InputWorkbooks";

        // Temporary folder to store generated PDFs
        string pdfFolder = Path.Combine(Path.GetTempPath(), "AsposePdfTemp");
        Directory.CreateDirectory(pdfFolder);

        // Path for the final ZIP archive
        string zipPath = @"C:\Output\WorkbooksArchive.zip";

        // Process each Excel file in the source folder
        foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Remove all slicers from every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards when removing items from a collection
                for (int i = sheet.Slicers.Count - 1; i >= 0; i--)
                {
                    sheet.Slicers.RemoveAt(i);
                }
            }

            // Save the modified workbook as PDF
            string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
            string pdfPath = Path.Combine(pdfFolder, pdfFileName);
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }

        // Create a ZIP archive containing all generated PDFs
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            foreach (string pdfFile in Directory.GetFiles(pdfFolder, "*.pdf"))
            {
                // Add each PDF to the archive with its file name
                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
            }
        }

        // Clean up temporary PDF files
        Directory.Delete(pdfFolder, true);
    }
}

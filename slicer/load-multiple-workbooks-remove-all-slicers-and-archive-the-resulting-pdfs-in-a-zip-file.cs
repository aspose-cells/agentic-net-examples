// Title: C# – Remove All Slicers from Multiple Excel Workbooks, Convert to PDF, and Zip the PDFs with Aspose.Cells
// Description: Loads a list of Excel files, deletes every slicer on each worksheet using the SlicerCollection.RemoveAt method, saves each workbook as a PDF, bundles all PDFs into a ZIP archive, and cleans up temporary files—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells slicer removal | C# batch Excel to PDF | zip multiple PDFs C# | remove slicers programmatically | Aspose.Cells PDF export | Excel workbook automation | temporary file cleanup
// Common Searches: how to delete all slicers in Excel using Aspose.Cells | batch convert Excel workbooks to PDF C# | zip generated PDFs from multiple workbooks | remove slicers before PDF export Aspose | C# code to archive PDFs in a zip file
// Developer Intent: The developer needs to process several Excel files, strip every slicer, export each workbook as a PDF, and deliver all PDFs as a single ZIP package.
// Use Cases: Automated report generation where slicer controls must be omitted before PDF distribution. | Compliance archiving of Excel reports as clean PDFs without interactive elements. | Providing end‑users a single downloadable ZIP containing PDF versions of multiple workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to remove all slicers from each worksheet of a workbook and then save it as a PDF. | Create a method that accepts a collection of Excel file paths, deletes slicers, converts each to PDF, and returns a ZIP file with all PDFs. | Explain best practices for deleting temporary PDF files after adding them to a ZipArchive when using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads a list of Excel files, deletes every slicer on each worksheet using the SlicerCollection.RemoveAt method, saves each workbook as a PDF, bundles all PDFs into a ZIP archive, and cleans up temporary files—all with Aspose.Cells for .NET.
class RemoveSlicersAndZipPdfs
{
    static void Main()
    {
        // Input Excel files (adjust paths as needed)
        List<string> excelFiles = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Folder to store intermediate PDFs
        string pdfFolder = Path.Combine(Path.GetTempPath(), "AsposePdfTemp");
        Directory.CreateDirectory(pdfFolder);

        // List to keep generated PDF file paths
        List<string> pdfFiles = new List<string>();

        foreach (string excelPath in excelFiles)
        {
            // Load workbook (uses provided constructor rule)
            Workbook workbook = new Workbook(excelPath);

            // Remove all slicers from every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;
                while (slicers.Count > 0)
                {
                    // RemoveAt method is the defined rule for deleting a slicer
                    slicers.RemoveAt(0);
                }
            }

            // Save workbook as PDF (uses provided Save method with SaveFormat)
            string pdfPath = Path.Combine(pdfFolder, Path.GetFileNameWithoutExtension(excelPath) + ".pdf");
            workbook.Save(pdfPath, SaveFormat.Pdf);
            pdfFiles.Add(pdfPath);

            // Dispose workbook resources
            workbook.Dispose();
        }

        // Create ZIP archive containing all PDFs
        string zipPath = "AllWorkbooks.pdf.zip";
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string pdfFile in pdfFiles)
            {
                // Add each PDF to the archive
                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
            }
        }

        // Clean up temporary PDF files
        foreach (string pdfFile in pdfFiles)
        {
            File.Delete(pdfFile);
        }

        // Optionally remove the temporary folder
        Directory.Delete(pdfFolder, true);

        Console.WriteLine($"PDFs archived successfully to '{zipPath}'.");
    }
}

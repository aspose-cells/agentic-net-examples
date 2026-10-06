// Title: Remove all slicers from multiple Excel workbooks, convert each workbook to PDF, and zip the PDFs using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a directory for *.xlsx files, loads each workbook with Aspose.Cells, deletes every slicer on all worksheets, saves the workbook as a PDF, and adds the PDF to a ZipArchive. | Create a reusable C# method that accepts input and output folder paths, removes slicers from every Excel file in the input folder, exports each to PDF, and returns the path of the generated zip file. | Generate a C# example that logs the name of each processed workbook, catches conversion errors, and optionally cleans up temporary PDF files after the zip archive is created.
// Common Searches: Aspose.Cells C# batch remove slicers from Excel workbooks and export to PDF | How to programmatically zip PDFs generated from multiple Excel files in .NET | C# example for deleting slicers in each worksheet before PDF conversion using Aspose.Cells | Automate conversion of a folder of .xlsx files to PDFs and archive them with ZipArchive
// Tags: batch slicer removal with Aspose.Cells | Excel workbook PDF export C# | create zip archive of PDFs .NET | automate workbook processing Aspose.Cells | delete worksheet slicers programmatically

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The sample loads every .xlsx file from a source folder, removes all slicers from each worksheet using Aspose.Cells, saves the modified workbook as a PDF in a temporary directory, and then packages all generated PDFs into a zip archive.
class WorkbookSlicerPdfArchiver
{
    static void Main()
    {
        // Folder containing the source workbooks
        string sourceFolder = @"C:\InputWorkbooks";

        // Temporary folder to store generated PDFs
        string pdfFolder = @"C:\TempPdfs";

        // Destination zip file path
        string zipPath = @"C:\Output\WorkbooksArchive.zip";

        // Ensure the PDF folder exists
        if (!Directory.Exists(pdfFolder))
            Directory.CreateDirectory(pdfFolder);

        // Process each workbook file in the source folder
        foreach (string workbookPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

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
            string pdfFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".pdf";
            string pdfPath = Path.Combine(pdfFolder, pdfFileName);
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }

        // Create a zip archive containing all generated PDFs
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            foreach (string pdfFile in Directory.GetFiles(pdfFolder, "*.pdf"))
            {
                // Add each PDF to the zip archive
                archive.CreateEntryFromFile(pdfFile, Path.GetFileName(pdfFile));
            }
        }

        // Optional: clean up temporary PDF files
        // Directory.Delete(pdfFolder, true);
    }
}

// Title: Batch convert Excel .xlsx files with WordArt to PDF using Aspose.Cells while preserving gradient fills (C#)
// AI Prompts: Write a C# console application that scans a directory for .xlsx workbooks, loads each with Aspose.Cells, and saves them as PDF files using PdfSaveOptions so that WordArt shapes retain their gradient fills. | Generate C# code that iterates over multiple Excel files, creates a Workbook for each, and exports to PDF with Aspose.Cells, ensuring shape rendering (including WordArt gradients) is maintained.
// Common Searches: asp.net batch convert excel files to pdf preserving wordart gradients | c# aspose.cells convert multiple .xlsx to pdf keep shape formatting | how to export excel workbooks with wordart to pdf using aspose.cells | pdfsaveoptions preserve gradient fills aspose.cells c# | process a folder of excel files and generate pdfs programmatically
// Tags: batch excel-to-pdf conversion using Aspose.Cells | wordart gradient preservation in PDF export | c# PdfSaveOptions shape rendering | iterate over xlsx files programmatically | export workbooks with embedded WordArt

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample program enumerates all .xlsx files in a given source folder, loads each workbook with Aspose.Cells inside a using block, and saves it as a PDF using PdfSaveOptions. The conversion runs in a batch, writing PDFs to an output directory while ensuring that WordArt objects, including their gradient fills, are rendered correctly in the resulting PDF files.
class BatchWordArtToPdf
{
    static void Main()
    {
        // Folder containing the source Excel files
        string sourceFolder = @"C:\Spreadsheets";
        // Folder where the resulting PDFs will be saved
        string outputFolder = @"C:\PdfOutput";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string excelPath in excelFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the workbook inside a using block for automatic disposal
                using (Workbook wb = new Workbook(excelPath))
                {
                    // Configure PDF save options (no need to set SaveFormat; it's inherent to PdfSaveOptions)
                    PdfSaveOptions pdfOptions = new PdfSaveOptions();

                    // Build the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    wb.Save(pdfPath, pdfOptions);
                }

                Console.WriteLine($"Converted '{excelPath}' to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}

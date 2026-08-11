// Title: Batch Convert Excel Files with WordArt Gradients to PDF using Aspose.Cells in C#
// Description: A C# console utility that scans a source directory for .xlsx workbooks, creates an output folder, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to PDF while preserving WordArt gradient fills. The script logs successes and failures and works for any number of spreadsheets in a single run.
// Keywords: Aspose.Cells | C# batch Excel to PDF | WordArt gradient conversion | ConversionUtility example | folder based PDF export | preserve graphics in PDF | console app Excel conversion | automated spreadsheet PDF generation
// Common Searches: batch convert Excel to PDF with WordArt using Aspose.Cells | C# code to preserve WordArt gradients when exporting Excel as PDF | Aspose.Cells ConversionUtility convert multiple workbooks | how to process a folder of .xlsx files to PDF in .NET | automate Excel PDF conversion preserving graphics
// Developer Intent: Convert every Excel workbook in a directory to PDF while keeping WordArt gradient rendering intact.
// Use Cases: Nightly job that archives report workbooks containing WordArt as PDF files. | Generating client‑ready PDFs from Excel templates without losing gradient colors. | Command‑line tool for bulk conversion of design‑heavy spreadsheets before distribution.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert Excel files to PDF and records any conversion errors. | Explain how to configure Aspose.Cells ConversionUtility to retain WordArt gradient fills in the PDF output. | Show how to extend the sample to walk subfolders recursively and convert all found .xlsx files to PDF.

using System;
using System.IO;
using Aspose.Cells.Utility; // Provides ConversionUtility

// A C# console utility that scans a source directory for .xlsx workbooks, creates an output folder, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to PDF while preserving WordArt gradient fills. The script logs successes and failures and works for any number of spreadsheets in a single run.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files with WordArt
        string sourceFolder = @"C:\Spreadsheets\Input";

        // Folder where the resulting PDFs will be saved
        string outputFolder = @"C:\Spreadsheets\Output";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the source folder
        foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Build the PDF file name based on the Excel file name
            string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            try
            {
                // Convert the Excel workbook (including WordArt gradients) to PDF
                ConversionUtility.Convert(excelPath, pdfPath);
                Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to convert '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}

// Title: C# Batch Processor to Hide Zeros, Apply 80% Zoom, Export Excel to PDF using Aspose.Cells
// Description: A console‑based C# utility that scans a given folder for .xlsx and .xls files, loads each workbook with Aspose.Cells, disables zero display, sets each worksheet's zoom to 80 %, and saves the result as a PDF in a target directory. Includes folder validation, automatic output creation, and robust error handling for seamless batch conversion.
// Keywords: Aspose.Cells batch PDF conversion | hide zero values Excel | worksheet zoom 80 percent | C# Excel to PDF automation | process multiple Excel files .NET | command line Excel PDF exporter | Aspose.Cells PDFSaveOptions | GitHub example Aspose.Cells
// Common Searches: how to hide zero values when exporting Excel to PDF with Aspose.Cells | batch convert a folder of .xlsx and .xls files to PDF in C# | set worksheet zoom to 80% before PDF export using Aspose.Cells | Aspose.Cells example for bulk Excel to PDF conversion | C# console app to process multiple workbooks and save as PDF
// Developer Intent: Automate the conversion of every Excel workbook in a directory to PDF while suppressing zero values and enforcing a uniform 80 % zoom on all worksheets.
// Use Cases: Nightly generation of client‑ready PDF reports from a batch of financial spreadsheets. | Command‑line tool for a document‑management pipeline that standardizes PDF appearance across dozens of workbooks. | Integration into CI/CD to verify that exported PDFs meet layout requirements before release.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert Excel files in a folder to PDF, hiding zero values and setting zoom to 80 %. | Show how to add structured logging and retry logic to the batch processor for resilient PDF conversion. | Demonstrate customizing PdfSaveOptions to embed fonts, set PDF/A compliance, and control image quality in the batch export.

using System;
using System.IO;
using Aspose.Cells;

// A console‑based C# utility that scans a given folder for .xlsx and .xls files, loads each workbook with Aspose.Cells, disables zero display, sets each worksheet's zoom to 80 %, and saves the result as a PDF in a target directory. Includes folder validation, automatic output creation, and robust error handling for seamless batch conversion.
public class WorkbookBatchProcessor
{
    // Processes all Excel files in the input folder, hides zero values,
    // sets worksheet zoom to 80%, and saves each workbook as a PDF.
    public void Process(string inputFolder, string outputFolder)
    {
        if (!Directory.Exists(inputFolder))
        {
            throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all Excel files (XLSX and XLS) in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");
        string[] oldExcelFiles = Directory.GetFiles(inputFolder, "*.xls");
        string[] allFiles = new string[excelFiles.Length + oldExcelFiles.Length];
        excelFiles.CopyTo(allFiles, 0);
        oldExcelFiles.CopyTo(allFiles, excelFiles.Length);

        foreach (string filePath in allFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply settings to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide zero values
                    sheet.DisplayZeros = false;

                    // Set zoom to 80%
                    sheet.PageSetup.Zoom = 80;
                }

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // (Optional) configure additional options here if needed

                // Determine output PDF file name
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"Converted to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}

public class Program
{
    // Entry point for the console application
    public static void Main(string[] args)
    {
        // Allow passing input and output folders via command‑line arguments
        string inputFolder = args.Length > 0 ? args[0] : @"C:\InputExcels";
        string outputFolder = args.Length > 1 ? args[1] : @"C:\OutputPDFs";

        try
        {
            var processor = new WorkbookBatchProcessor();
            processor.Process(inputFolder, outputFolder);
            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}

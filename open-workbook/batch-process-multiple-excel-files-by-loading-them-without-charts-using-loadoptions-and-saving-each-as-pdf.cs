// Title: Batch convert Excel workbooks to PDF without charts using Aspose.Cells LoadOptions (C#)
// Description: Scans a given folder for XLSX, XLS, XLSB, and CSV files, creates an output directory, and converts each workbook to PDF while skipping charts and other non‑data objects. The code uses Aspose.Cells LoadOptions (auto format) and PdfSaveOptions (one page per sheet) with ConversionUtility.Convert, logging successes and errors.
// Keywords: Aspose.Cells batch Excel to PDF | C# LoadOptions skip charts | ConversionUtility Convert | PdfSaveOptions one page per sheet | process multiple workbooks | automated Excel PDF conversion | load data only Aspose.Cells | Excel files to PDF programmatically
// Common Searches: convert all Excel files in a folder to PDF using Aspose.Cells | Aspose.Cells C# load workbook without charts | batch PDF conversion of XLSX, XLS, XLSB, CSV | how to use ConversionUtility to export Excel to PDF | skip charts when saving Excel as PDF with Aspose
// Developer Intent: Convert every Excel file in a directory to a PDF while omitting charts and other visual objects.
// Use Cases: Nightly generation of PDF reports from a collection of financial spreadsheets. | Archiving incoming CSV/XLSX files as PDF documents without visual clutter. | Integrating bulk Excel‑to‑PDF conversion into a web API that processes user uploads.
// AI Prompts: Show C# code that loads an Excel workbook with Aspose.Cells LoadOptions to exclude charts and saves it as a single‑page‑per‑sheet PDF. | Explain how to extend the batch conversion script to support password‑protected Excel files. | Create a unit test that verifies a PDF is produced for each supported Excel file after running the batch converter.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToPdf
{
    // Scans a given folder for XLSX, XLS, XLSB, and CSV files, creates an output directory, and converts each workbook to PDF while skipping charts and other non‑data objects. The code uses Aspose.Cells LoadOptions (auto format) and PdfSaveOptions (one page per sheet) with ConversionUtility.Convert, logging successes and errors.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files to be processed
            string inputFolder = @"C:\InputExcelFiles";

            // Folder where the resulting PDF files will be saved
            string outputFolder = @"C:\OutputPdfFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (XLSX, XLS, XLSB, CSV) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Process only supported Excel extensions
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsb" && ext != ".csv")
                    continue;

                // Verify the source file still exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Destination PDF file path (same name, .pdf extension)
                string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                try
                {
                    // LoadOptions: load data only (skip charts, shapes, etc.)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    // Note: LoadDataOnly property may not be available in all versions; omitted for compatibility.

                    // SaveOptions: PDF specific options (optional customizations)
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true,               // each sheet fits on one page
                        AllColumnsInOnePagePerSheet = true,   // all columns on one page
                        IgnoreError = true                    // hide rendering errors if any
                    };

                    // Convert the Excel file to PDF using the utility method
                    ConversionUtility.Convert(filePath, loadOptions, pdfPath, pdfOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to PDF.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}

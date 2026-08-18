// Title: Batch convert Excel (XLSX, XLS, CSV) to PDF with custom options using Aspose.Cells for .NET
// Description: Creates an output folder, loops through a list of Excel files, validates each path, builds a matching PDF name, and uses Aspose.Cells LoadOptions together with PdfSaveOptions (OnePagePerSheet, AllColumnsInOnePagePerSheet, PDF/A‑1b compliance, formula recalculation, font embedding) to convert every workbook to PDF via ConversionUtility.Convert. Successes and errors are logged to the console.
// Keywords: Aspose.Cells batch conversion | Excel to PDF C# | PdfSaveOptions OnePagePerSheet | PDF/A-1b Aspose.Cells | ConversionUtility.Convert example | multiple Excel files to PDF | C# Excel PDF batch processing | embed fonts Aspose PDF
// Common Searches: How to batch convert XLSX, XLS, CSV to PDF with Aspose.Cells | Set per‑file PDF options in Aspose.Cells C# | Convert many Excel workbooks to PDF/A‑1b using .NET | Aspose.Cells error handling for missing Excel files | C# code for converting a list of spreadsheets to PDF
// Developer Intent: Convert a collection of Excel workbooks to PDF in a single run while applying individual PdfSaveOptions to each file.
// Use Cases: Automate daily generation of PDF reports from multiple Excel exports. | Archive financial spreadsheets as PDF/A‑1b compliant documents for regulatory purposes. | Produce portable PDFs from CSV data sets with formulas recalculated and fonts embedded.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert an array of Excel file paths to PDF, applying custom PdfSaveOptions such as OnePagePerSheet and PDF/A‑1b compliance, with error handling for missing files. | Explain how to vary PdfSaveOptions per file based on filename patterns in a batch conversion scenario. | Show how to log conversion results to a CSV file instead of the console while using Aspose.Cells ConversionUtility.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchExcelToPdf
{
    // Creates an output folder, loops through a list of Excel files, validates each path, builds a matching PDF name, and uses Aspose.Cells LoadOptions together with PdfSaveOptions (OnePagePerSheet, AllColumnsInOnePagePerSheet, PDF/A‑1b compliance, formula recalculation, font embedding) to convert every workbook to PDF via ConversionUtility.Convert. Successes and errors are logged to the console.
    public class Converter
    {
        /// <param name="sourceFiles">Full paths of the Excel files to convert.</param>
        /// <param name="outputFolder">Folder where the resulting PDFs will be saved.</param>
        public static void ConvertBatch(string[] sourceFiles, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            foreach (string sourcePath in sourceFiles)
            {
                // Validate source file existence.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    continue;
                }

                // Determine the output PDF file name (same base name, .pdf extension).
                string outputPath = Path.Combine(outputFolder,
                    Path.GetFileNameWithoutExtension(sourcePath) + ".pdf");

                // Create LoadOptions based on the source file format.
                // For simplicity we let Aspose infer the format, but we can specify it explicitly.
                LoadOptions loadOptions = new LoadOptions();

                // Create PDF save options with custom settings for this file.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example customizations – adjust as needed per file.
                    OnePagePerSheet = true,                     // Fit each sheet onto a single page.
                    AllColumnsInOnePagePerSheet = true,        // Fit all columns onto one page.
                    Compliance = PdfCompliance.PdfA1b,         // PDF/A‑1b compliance.
                    CalculateFormula = true,                   // Re‑calculate formulas before saving.
                    EmbedStandardWindowsFonts = true           // Embed fonts for better portability.
                };

                try
                {
                    // Perform the conversion using the utility method that accepts
                    // load and save options.
                    ConversionUtility.Convert(sourcePath, loadOptions, outputPath, pdfOptions);
                    Console.WriteLine($"Converted: {sourcePath} → {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }
        }

        // Example usage.
        public static void Main()
        {
            // Define the Excel files to process.
            string[] filesToConvert = new[]
            {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xls",
                @"C:\Data\Report3.csv"
            };

            // Destination folder for PDFs.
            string pdfFolder = @"C:\Data\PdfOutput";

            ConvertBatch(filesToConvert, pdfFolder);
        }
    }
}

// Title: Batch convert Excel (.xlsx, .xls) files to PDF with per‑file PdfSaveOptions in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells batch convert a list of Excel files to PDF with different OnePagePerSheet settings | How to apply individual PdfSaveOptions for each workbook when exporting to PDF in .NET | Example code for converting multiple .xlsx and .xls files to PDF using Aspose.Cells with custom PDF options per file | Aspose.Cells convert Excel to PDF in a loop with per‑file PDF export settings | Batch export Excel workbooks to PDF with custom OnePagePerSheet option using Aspose.Cells for .NET
// Tags: Aspose.Cells batch Excel to PDF conversion | PdfSaveOptions per-file customization | OnePagePerSheet setting Aspose.Cells PDF export | C# convert .xls and .xlsx to PDF with Aspose.Cells | dictionary based PDF options for Excel workbooks

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToPdf
{
    // The sample iterates a predefined list of .xlsx and .xls files, loads each workbook with Aspose.Cells, looks up a corresponding PdfSaveOptions entry from a dictionary (or falls back to defaults), and saves the workbook as a PDF in an output folder, allowing per‑file settings such as OnePagePerSheet while handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // List of Excel files to convert
            List<string> excelFiles = new List<string>
            {
                @"C:\Input\Report1.xlsx",
                @"C:\Input\Report2.xls",
                @"C:\Input\Report3.xlsx"
            };

            // Per‑file custom PDF options (only OnePagePerSheet is set; other settings are defaults)
            Dictionary<string, PdfSaveOptions> pdfOptions = new Dictionary<string, PdfSaveOptions>
            {
                {
                    @"C:\Input\Report1.xlsx",
                    new PdfSaveOptions
                    {
                        OnePagePerSheet = false
                    }
                },
                {
                    @"C:\Input\Report2.xls",
                    new PdfSaveOptions
                    {
                        OnePagePerSheet = true
                    }
                },
                {
                    @"C:\Input\Report3.xlsx",
                    new PdfSaveOptions() // default options
                }
            };

            // Output folder for PDFs
            string outputFolder = @"C:\Output\PDFs";
            Directory.CreateDirectory(outputFolder);

            foreach (string excelPath in excelFiles)
            {
                // Verify the source file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Source file not found: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Retrieve custom options for the current file (or use defaults)
                    if (!pdfOptions.TryGetValue(excelPath, out PdfSaveOptions options) || options == null)
                    {
                        options = new PdfSaveOptions(); // fallback to default options
                    }

                    // Build output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save workbook as PDF with the specified options
                    workbook.Save(pdfPath, options);

                    Console.WriteLine($"Converted '{excelPath}' to PDF.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}

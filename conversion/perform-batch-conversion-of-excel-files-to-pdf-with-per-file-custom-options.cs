// Title: Batch Convert Excel & CSV to PDF with Individual PdfSaveOptions – Aspose.Cells for .NET
// Description: C# sample that iterates a list of Excel/CSV files, detects each format, applies a custom PdfSaveOptions (one‑page‑per‑sheet, all‑columns‑in‑one‑page, PDF/A‑1b, watermark, image resampling) and converts them to PDF using Aspose.Cells ConversionUtility, with error handling and file‑existence checks.
// Keywords: Aspose.Cells batch conversion | Excel to PDF C# | CSV to PDF Aspose | PdfSaveOptions per file | one page per sheet Aspose | PDF/A compliance Aspose.Cells | watermark PDF Aspose.Cells | image resample PDF Aspose | ConversionUtility Convert method | load format auto detection
// Common Searches: batch convert multiple Excel files to PDF with different settings Aspose.Cells | apply watermark during Excel to PDF conversion C# | set PDF/A‑1b for specific Excel files in batch conversion | one page per sheet option for batch Excel to PDF | convert CSV to PDF with image resampling using Aspose.Cells
// Developer Intent: Convert a mixed collection of .xlsx, .xls, and .csv workbooks to PDF, assigning a unique PdfSaveOptions configuration to each file.
// Use Cases: Create separate PDF reports where each workbook requires a distinct layout—single‑page‑per‑sheet, all columns on one page, or PDF/A compliance. | Add a semi‑transparent diagonal watermark and lower image resolution for sensitive CSV data before generating PDFs. | Automatically detect CSV files, select the proper LoadFormat, and process a heterogeneous batch with individual conversion options.
// AI Prompts: Generate C# code that adds a diagonal semi‑transparent watermark to PDFs produced from Excel files with Aspose.Cells. | Show how to extend the batch conversion loop to record conversion duration for each file and export a CSV summary. | Explain how to configure PdfSaveOptions to embed fonts and enable PDF/A‑2b compliance for selected files in a batch process.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchExcelToPdf
{
    // C# sample that iterates a list of Excel/CSV files, detects each format, applies a custom PdfSaveOptions (one‑page‑per‑sheet, all‑columns‑in‑one‑page, PDF/A‑1b, watermark, image resampling) and converts them to PDF using Aspose.Cells ConversionUtility, with error handling and file‑existence checks.
    class Program
    {
        static void Main()
        {
            // Define source Excel files and corresponding PDF output paths
            var files = new List<(string source, string dest, Action<PdfSaveOptions> configure)>
            {
                // Example 1: simple conversion, one page per sheet
                ("Input1.xlsx", "Output1.pdf", options =>
                {
                    options.OnePagePerSheet = true;
                }),

                // Example 2: fit all columns on one page, use PDF/A-1b compliance
                ("Input2.xls", "Output2.pdf", options =>
                {
                    options.AllColumnsInOnePagePerSheet = true;
                    options.Compliance = PdfCompliance.PdfA1b;
                }),

                // Example 3: custom watermark and image resampling
                ("Input3.csv", "Output3.pdf", options =>
                {
                    options.Watermark = new RenderingWatermark("CONFIDENTIAL",
                        new RenderingFont("Arial", 48) { Color = System.Drawing.Color.Red, Bold = true })
                    {
                        Opacity = 0.3f,
                        Rotation = 45
                    };
                    options.SetImageResample(150, 80); // 150 PPI, 80% JPEG quality
                })
            };

            foreach (var (source, dest, configure) in files)
            {
                // Verify that the source file exists
                if (!File.Exists(source))
                {
                    Console.WriteLine($"Source file '{source}' not found. Skipping conversion.");
                    continue;
                }

                // Determine load format based on file extension
                LoadFormat loadFormat = source.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                    ? LoadFormat.Csv
                    : LoadFormat.Auto; // Auto detects format for other extensions

                // Create load options (if needed)
                var loadOptions = new LoadOptions(loadFormat);

                // Create PDF save options and apply per‑file custom configuration
                var pdfOptions = new PdfSaveOptions();
                configure?.Invoke(pdfOptions);

                try
                {
                    // Perform conversion using the provided utility method
                    ConversionUtility.Convert(source, loadOptions, dest, pdfOptions);
                    Console.WriteLine($"Converted '{source}' to PDF '{dest}' with custom options.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{source}' to PDF: {ex.Message}");
                }
            }
        }
    }
}

// Title: Save a merged Aspose.Cells workbook to XLSX and optionally export a PDF report in C#
// Common Searches: C# code to save an Aspose.Cells merged workbook as XLSX with optional PDF output | How to conditionally export a workbook to PDF using Aspose.Cells SaveFormat in C# | Aspose.Cells save merged workbook to XLSX and generate PDF only when a path is provided | Example of using Aspose.Cells SaveFormat.Xlsx and SaveFormat.Pdf together in C# | Error handling when saving Aspose.Cells workbook to XLSX and PDF files
// Tags: Aspose.Cells save workbook as xlsx | Aspose.Cells export workbook to pdf | C# conditional workbook export | Aspose.Cells SaveFormat usage | merged workbook saving Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The WorkbookExporter class validates the merged Workbook and the XLSX output path, saves the workbook as an XLSX file using SaveFormat.Xlsx, and if a non‑empty PDF path is supplied, also saves it as a PDF with SaveFormat.Pdf, handling any exceptions that may occur.
    public class WorkbookExporter
    {
        /// <param name="mergedWorkbook">The merged Workbook instance to be saved.</param>
        /// <param name="xlsxPath">File path for the XLSX output (e.g., "MergedWorkbook.xlsx").</param>
        /// <param name="pdfPath">Optional file path for the PDF output. Pass null or empty string to skip PDF export.</param>
        public static void SaveWorkbook(Workbook mergedWorkbook, string xlsxPath, string pdfPath = null)
        {
            if (mergedWorkbook == null)
                throw new ArgumentNullException(nameof(mergedWorkbook));

            if (string.IsNullOrWhiteSpace(xlsxPath))
                throw new ArgumentException("XLSX path must be provided.", nameof(xlsxPath));

            try
            {
                // Save the workbook as XLSX
                mergedWorkbook.Save(xlsxPath, SaveFormat.Xlsx);

                // If a PDF path is supplied, export the workbook to PDF as well
                if (!string.IsNullOrWhiteSpace(pdfPath))
                {
                    mergedWorkbook.Save(pdfPath, SaveFormat.Pdf);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Determine output paths
            string outputDir = Directory.GetCurrentDirectory();
            string xlsxPath = Path.Combine(outputDir, "MergedWorkbook.xlsx");
            string pdfPath = Path.Combine(outputDir, "MergedWorkbook.pdf");

            Workbook workbook = null;

            // Optional: load a template workbook if provided as first argument
            if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                string templatePath = args[0];
                if (File.Exists(templatePath))
                {
                    try
                    {
                        workbook = new Workbook(templatePath);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to load template '{templatePath}': {ex.Message}");
                    }
                }
                else
                {
                    Console.Error.WriteLine($"Template file not found: {templatePath}");
                }
            }

            // If no workbook loaded, create a new one with sample data
            if (workbook == null)
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sample";
                sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");
            }

            // Save workbook using exporter
            try
            {
                WorkbookExporter.SaveWorkbook(workbook, xlsxPath, pdfPath);
                Console.WriteLine($"Workbook saved successfully to '{xlsxPath}' and '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}

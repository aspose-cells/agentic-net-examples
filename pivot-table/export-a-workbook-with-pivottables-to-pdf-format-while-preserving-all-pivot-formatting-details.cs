// Title: Export an Excel workbook containing PivotTables to PDF with full formatting using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook with pivot tables to PDF preserving layout | How to keep pivot table formatting when saving Excel as PDF using Aspose.Cells | PdfSaveOptions settings to retain pivot table appearance in C# | Export Excel file with multiple pivot tables to PDF without scaling using Aspose.Cells | Aspose.Cells PDF export all columns on one page pivot tables C#
// Tags: Aspose.Cells PdfSaveOptions pivot layout preservation | C# export Excel pivot tables to PDF | Excel to PDF with pivot formatting using Aspose | AllColumnsInOnePagePerSheet option for pivot tables | OnePagePerSheet false PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace PivotTablePdfExportApp
{
    // The sample loads an existing Excel workbook (input.xlsx) that contains PivotTables, configures PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = true) to keep the original layout, and saves the workbook as a PDF (output.pdf) while handling missing files and runtime errors.
    class PivotTablePdfExport
    {
        static void Main()
        {
            // Path to the source workbook containing PivotTables
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "output.pdf";

            try
            {
                // Ensure the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve the layout as shown in Excel
                    OnePagePerSheet = false,
                    // Ensure all columns fit without scaling down
                    AllColumnsInOnePagePerSheet = true
                };

                // Export the workbook (including PivotTables) to PDF
                workbook.Save(pdfFile, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {pdfFile}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

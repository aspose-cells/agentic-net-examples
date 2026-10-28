// Title: Export an Excel workbook containing PivotTables to PDF while retaining all pivot formatting with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook with pivot tables to PDF preserving formatting | how to keep pivot table styles when converting Excel to PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions settings for pivot table layout in .NET | refresh pivot data before saving Excel as PDF with Aspose.Cells C# | save Excel file as PDF without losing pivot table formatting Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions preserve pivot formatting | C# export Excel to PDF with pivot tables | refresh pivot data before PDF conversion Aspose.Cells | disable AllColumnsInOnePagePerSheet Aspose.Cells | export workbook to PDF retaining layout Aspose.Cells | pivot table formatting preservation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the input file, loads the workbook with Aspose.Cells, recalculates formulas to update PivotTable data, configures PdfSaveOptions (AllColumnsInOnePagePerSheet = false, OnePagePerSheet = false) to maintain the original sheet layout, and saves the workbook as a PDF while automatically preserving all PivotTable formatting. It includes handling for Aspose.Cells‑specific and generic exceptions.
class ExportPivotToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains PivotTables.
            Workbook workbook = new Workbook(inputPath);

            // Refresh all formulas (including PivotTable data) to ensure the latest values are used.
            workbook.CalculateFormula();

            // Configure PDF save options to keep the original layout and formatting.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force all columns onto a single page; keep the sheet's original width.
                AllColumnsInOnePagePerSheet = false,
                // Do not force each sheet onto a separate PDF page.
                OnePagePerSheet = false
                // PivotTable formatting is preserved automatically during PDF export.
            };

            // Export the workbook to PDF while preserving PivotTable formatting.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password-protected files).
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

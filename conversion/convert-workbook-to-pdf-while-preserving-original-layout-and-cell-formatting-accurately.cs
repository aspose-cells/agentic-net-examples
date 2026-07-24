// Title: C# – Convert Excel Workbook to PDF with Layout & Formatting Preservation using Aspose.Cells
// Description: Shows how to turn an .xlsx into a PDF in C# with Aspose.Cells while keeping cell styles, column widths, and document structure. Includes a quick ConversionUtility call and a detailed Workbook.Save using PdfSaveOptions (ExportDocumentStructure, CheckWorkbookDefaultFont, OnePagePerSheet = false).
// Keywords: Aspose.Cells PDF conversion C# | Excel to PDF preserving layout | PdfSaveOptions ExportDocumentStructure | ConversionUtility Convert example | Workbook.Save PDF options | calculate formulas before export | C# Excel PDF export
// Common Searches: Aspose.Cells convert xlsx to pdf preserving formatting | C# PdfSaveOptions keep original layout Excel | Difference between ConversionUtility.Convert and Workbook.Save for PDF | How to export Excel as multi‑page PDF with Aspose.Cells | Preserve column widths when saving Excel to PDF C#
// Developer Intent: Create a PDF from an Excel workbook that exactly mirrors the original worksheet’s layout and formatting.
// Use Cases: Rapid conversion when default PDF layout is sufficient – use ConversionUtility.Convert. | Fine‑tuned export that retains document hierarchy, default fonts, and multi‑page sheet layout – use Workbook.Save with customized PdfSaveOptions. | Ensure formulas are up‑to‑date before export by calling Workbook.CalculateFormula.
// AI Prompts: Generate C# code to convert an .xlsx to PDF with Aspose.Cells, keeping column widths and row heights unchanged. | Provide a snippet that uses PdfSaveOptions to export an Excel workbook to a multi‑page PDF while preserving document structure and default fonts. | Explain the impact of OnePagePerSheet and AllColumnsInOnePagePerSheet settings on PDF layout in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExample
{
    // Shows how to turn an .xlsx into a PDF in C# with Aspose.Cells while keeping cell styles, column widths, and document structure. Includes a quick ConversionUtility call and a detailed Workbook.Save using PdfSaveOptions (ExportDocumentStructure, CheckWorkbookDefaultFont, OnePagePerSheet = false).
    class WorkbookToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // 1. Quick conversion using the utility method
                ConversionUtility.Convert(sourcePath, destPath);
                Console.WriteLine("ConversionUtility: Workbook converted to PDF.");

                // 2. Precise conversion with full control over PDF options
                Workbook workbook = new Workbook(sourcePath);
                workbook.CalculateFormula(); // Ensure formulas are up‑to‑date

                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,
                    CheckWorkbookDefaultFont = true,
                    OnePagePerSheet = false,
                    AllColumnsInOnePagePerSheet = false
                };

                workbook.Save(destPath, pdfOptions);
                Console.WriteLine("Workbook.Save: Workbook saved to PDF with layout preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WorkbookToPdfConverter.Run();
        }
    }
}

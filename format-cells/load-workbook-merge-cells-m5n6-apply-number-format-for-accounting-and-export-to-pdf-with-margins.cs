// Title: Merge M5:N6, apply accounting format, set margins, and export to PDF with Aspose.Cells (C#)
// Description: C# example that loads an Excel workbook, merges the range M5:N6, applies the built‑in accounting number format (ID 44) to the merged cells, configures 0.5‑inch page margins, and saves the worksheet as a PDF using Aspose.Cells PdfSaveOptions with document structure preserved.
// Keywords: Aspose.Cells merge cells C# | accounting number format 44 Aspose | set page margins Aspose.Cells | export Excel to PDF .NET | PdfSaveOptions document structure | C# Excel to PDF conversion | Aspose.Cells formatting merged range
// Common Searches: how to merge cells and apply accounting format with Aspose.Cells C# | set custom margins when exporting Excel to PDF using Aspose.Cells | apply built‑in number format 44 to a merged range in Aspose.Cells | export workbook to PDF preserving document structure Aspose.Cells .NET | Aspose.Cells C# example for PDF export with page margins
// Developer Intent: Merge cells M5:N6, format them with the accounting style, define 0.5‑inch margins, and generate a PDF from the workbook.
// Use Cases: Financial statements where a total row spans columns M‑N and must appear in accounting format in the printed PDF. | Invoices that require a merged amount cell formatted as accounting and a consistent margin layout for printing. | Standardized report generation that enforces half‑inch margins and retains cell styling when converting Excel to PDF.
// AI Prompts: Generate C# code to merge M5:N6, apply accounting format 44, set 0.5‑inch margins, and export to PDF with Aspose.Cells. | Show how to use PdfSaveOptions to keep document structure while saving an Excel worksheet as PDF in .NET. | Explain the steps to create a StyleFlag that applies all formatting to a merged range before PDF conversion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // C# example that loads an Excel workbook, merges the range M5:N6, applies the built‑in accounting number format (ID 44) to the merged cells, configures 0.5‑inch page margins, and saves the worksheet as a PDF using Aspose.Cells PdfSaveOptions with document structure preserved.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge cells M5:N6 (zero‑based: row 4, column 12, spanning 2 rows and 2 columns)
                cells.Merge(4, 12, 2, 2);

                // Apply accounting number format (built‑in format 44) to the merged area
                Style accountingStyle = cells["M5"].GetStyle();
                accountingStyle.Number = 44; // Accounting format

                // Apply the style to the merged range
                AsposeRange mergedRange = worksheet.Cells.CreateRange("M5:N6");
                mergedRange.ApplyStyle(accountingStyle, new StyleFlag { All = true });

                // Set page margins (in inches)
                worksheet.PageSetup.LeftMargin = 0.5;
                worksheet.PageSetup.RightMargin = 0.5;
                worksheet.PageSetup.TopMargin = 0.5;
                worksheet.PageSetup.BottomMargin = 0.5;

                // Save the workbook as PDF with desired options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true // optional: keep document structure
                };
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

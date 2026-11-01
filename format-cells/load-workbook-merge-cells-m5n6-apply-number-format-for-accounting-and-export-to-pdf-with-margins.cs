// Title: C# – Merge cells M5:N6, apply accounting number format, set PDF page margins, and export workbook to PDF using Aspose.Cells
// Common Searches: how to merge cells M5:N6 with Aspose.Cells in C# before PDF export | apply accounting number format to a merged range using Aspose.Cells .NET | set custom PDF page margins when saving an Excel workbook with Aspose.Cells | export a worksheet to PDF with OnePagePerSheet option in Aspose.Cells C# | Aspose.Cells example for merging cells and formatting numbers for PDF output
// Tags: merge cells Aspose.Cells C# | accounting number format style Aspose.Cells | set PDF page margins Aspose.Cells | PdfSaveOptions OnePagePerSheet Aspose.Cells | export workbook to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# program loads an existing Excel file (or creates a new workbook if missing), merges cells M5:N6 on the first worksheet, applies a custom accounting number format to the merged range, sets all PDF page margins to 0.5 inches, and saves the workbook as a PDF using PdfSaveOptions with OnePagePerSheet enabled.
class Program
{
    static void Main()
    {
        // Paths for input Excel and output PDF
        string inputPath = "input.xlsx";
        string outputPdf = "output.pdf";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least one worksheet exists
                workbook.Worksheets.Add();
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Merge cells M5:N6 (zero‑based indices)
            // Column M = 12, Column N = 13
            // Row 5   = 4,   Row 6   = 5
            // -------------------------------------------------
            sheet.Cells.Merge(12, 4, 2, 2); // startColumn, startRow, totalColumns, totalRows

            // Create a range that covers the merged cells
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("M5:N6");

            // -------------------------------------------------
            // Apply Accounting number format to the merged range
            // -------------------------------------------------
            Style accountingStyle = workbook.CreateStyle();
            accountingStyle.Custom = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            mergedRange.ApplyStyle(accountingStyle, flag);

            // -------------------------------------------------
            // Set page margins (in inches) for PDF export
            // -------------------------------------------------
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;

            // Export the workbook to PDF with custom options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            workbook.Save(outputPdf, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

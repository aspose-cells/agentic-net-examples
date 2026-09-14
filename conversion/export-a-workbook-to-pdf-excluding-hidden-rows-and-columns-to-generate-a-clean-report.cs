// Title: Export an Aspose.Cells workbook to PDF in C# while excluding hidden rows and columns for a clean report
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as a PDF, automatically skipping any concealed rows and columns. | Show how to configure PdfSaveOptions in Aspose.Cells so that hidden worksheet elements are not rendered in the exported PDF. | Provide a complete example that hides specific rows and columns, then creates a PDF report without those hidden cells.
// Common Searches: Aspose.Cells C# export to PDF without hidden rows | How to ignore hidden columns when converting Excel to PDF with Aspose.Cells | C# PdfSaveOptions to exclude hidden worksheet elements in PDF conversion | Generate clean PDF report from Excel workbook while omitting hidden cells using Aspose.Cells | Aspose.Cells ignore blank pages and hidden rows during PDF export C#
// Tags: Aspose.Cells PDF export invisible rows | C# hide worksheet columns before PDF conversion | PdfSaveOptions ignore non‑visible elements | clean PDF report from Excel workbook | Aspose.Cells exclude concealed cells PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The program creates a new workbook, populates sample data, hides row 2 and column B, configures PdfSaveOptions (including IgnoringBlank pages), and saves the file as CleanReport.pdf where the concealed rows and columns are omitted from the generated PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");
            sheet.Cells["C2"].PutValue("Data3");
            sheet.Cells["A3"].PutValue("Data4");
            sheet.Cells["B3"].PutValue("Data5");
            sheet.Cells["C3"].PutValue("Data6");

            // Hide row 2 (index 1) and column B (index 1)
            sheet.Cells.HideRow(1);
            sheet.Cells.HideColumn(1);

            // Set PDF save options (optional: ignore blank pages)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PrintingPageType = PrintingPageType.IgnoreBlank,
                // Hidden rows/columns are not rendered by default, no extra flag needed
                // Additional options can be set here if required
            };

            // Save the workbook to PDF; hidden rows/columns will be excluded
            workbook.Save("CleanReport.pdf", pdfOptions);
        }
    }
}

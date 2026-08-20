// Title: C# – Export Excel to PDF while Skipping Empty Worksheets with Aspose.Cells
// Description: This example shows how to detect worksheets that contain no data and exclude them from PDF conversion. It uses MaxDataRow/MaxDataColumn to find populated sheets, builds a SheetSet, and configures PdfSaveOptions (OutputBlankPageWhenNothingToPrint = false, PrintingPageType = IgnoreBlank) to generate a clean PDF without blank pages.
// Keywords: Aspose.Cells PDF export C# | skip empty worksheets Aspose | PdfSaveOptions OutputBlankPageWhenNothingToPrint | PrintingPageType IgnoreBlank | SheetSet non‑empty sheets | remove blank pages Excel to PDF | global | USA | India
// Common Searches: How to prevent blank pages when saving Excel as PDF with Aspose.Cells | Aspose.Cells skip empty sheets during PDF conversion | C# code to export only populated worksheets to PDF | Ignore blank worksheets Aspose.Cells PDF export | PdfSaveOptions settings for removing empty pages
// Developer Intent: Exclude worksheets that have no content so the resulting PDF contains only pages with actual data.
// Use Cases: Create PDF reports from workbooks that may contain placeholder or template sheets, ensuring the final document shows only relevant data. | Batch‑process a folder of Excel files, automatically omitting blank worksheets to reduce PDF size and improve readability. | Build a web API that receives an Excel file and returns a PDF that filters out any completely empty sheets.
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to PDF, omitting worksheets with no populated cells. | Show how to configure PdfSaveOptions to disable blank‑page generation and specify a SheetSet of non‑empty worksheet indexes. | Explain the method for detecting an empty worksheet in Aspose.Cells using MaxDataRow and MaxDataColumn.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to detect worksheets that contain no data and exclude them from PDF conversion. It uses MaxDataRow/MaxDataColumn to find populated sheets, builds a SheetSet, and configures PdfSaveOptions (OutputBlankPageWhenNothingToPrint = false, PrintingPageType = IgnoreBlank) to generate a clean PDF without blank pages.
class GeneratePdfSkippingEmptySheets
{
    static void Main()
    {
        // Create a new workbook (replace with loading if needed)
        Workbook workbook = new Workbook();

        // Example data: first sheet has content, second sheet is empty
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
        workbook.Worksheets.Add("EmptySheet");

        // Collect indexes of worksheets that contain at least one cell with data
        List<int> printableSheetIndexes = new List<int>();
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            // MaxDataRow/MaxDataColumn are -1 when the sheet has no data
            if (sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0)
            {
                printableSheetIndexes.Add(i);
            }
        }

        // If no printable sheets exist, exit early
        if (printableSheetIndexes.Count == 0)
        {
            Console.WriteLine("Workbook contains no printable content.");
            return;
        }

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Prevent generation of a blank page when a sheet has nothing to print
            OutputBlankPageWhenNothingToPrint = false,
            // Omit pages that are completely blank after rendering
            PrintingPageType = PrintingPageType.IgnoreBlank,
            // Render only the identified non‑empty sheets
            SheetSet = new SheetSet(printableSheetIndexes.ToArray())
        };

        // Save the workbook to PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
        Console.WriteLine("PDF saved successfully without empty worksheets.");
    }
}

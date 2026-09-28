// Title: Hide rows 50‑55 in an Excel worksheet and export the visible area to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file, hides rows 50 through 55 using Aspose.Cells, and saves the result as a PDF. | Provide a snippet showing how to call Cells.HideRows together with PdfSaveOptions to create a PDF that omits hidden rows. | Demonstrate how to hide a specific row range in a worksheet and export only the visible rows to PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells hide specific rows before converting to PDF in C# | C# code to hide rows 50-55 in Excel and export to PDF using Aspose.Cells | How to prevent hidden rows from appearing in PDF output with Aspose.Cells .NET | Using Cells.HideRows and PdfSaveOptions to generate PDF without hidden rows
// Tags: cells.hiderows pdf export | aspose.cells hide rows c# | pdfsaveoptions exclude hidden rows | excel to pdf conversion aspose.cells | worksheet row visibility aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Loads input.xlsx, hides rows 50‑55 on the first worksheet via Cells.HideRows, and saves the workbook as output.pdf; the PDF contains only the visible rows.
    class HideRowsAndExportPdf
    {
        static void Main()
        {
            // Load an existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 50 to 55 (zero‑based index 49, total 6 rows)
            worksheet.Cells.HideRows(49, 6);

            // Set PDF save options (default options are sufficient for hidden rows)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; hidden rows will not appear in the output
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}

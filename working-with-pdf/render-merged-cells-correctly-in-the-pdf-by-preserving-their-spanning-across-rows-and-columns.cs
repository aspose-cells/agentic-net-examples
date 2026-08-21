// Title: C# – Preserve Merged Cells When Converting Excel to PDF with Aspose.Cells
// Description: Demonstrates how to merge A1:B1, apply center‑aligned bold styling, and export a workbook to PDF while keeping the merged layout intact using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# PDF merge cells | preserve merged cells PDF | Excel to PDF Aspose .NET | merged header PDF export | Aspose.Cells SaveFormat.Pdf
// Common Searches: keep merged cells when saving Excel as PDF Aspose | Aspose.Cells .NET merged cell PDF rendering | how to export merged header to PDF C# | Aspose.Cells PDF options for merged ranges
// Developer Intent: Generate a PDF from an Excel workbook where any merged ranges remain merged and retain their formatting.
// Use Cases: Create a multi‑column report title that appears correctly in the PDF. | Export invoices with merged title cells without losing layout. | Produce dashboards where merged header cells keep center alignment and bold styling in the final PDF.
// AI Prompts: Show C# code to merge cells A1:B1, style them, and save the workbook as PDF with Aspose.Cells. | Explain which PdfSaveOptions are required to keep merged cells intact during PDF conversion. | Provide a step‑by‑step guide for preserving merged cell formatting when exporting Excel to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PDF rendering options if needed

// Demonstrates how to merge A1:B1, apply center‑aligned bold styling, and export a workbook to PDF while keeping the merged layout intact using Aspose.Cells for .NET.
class RenderMergedCellsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data
        cells["A1"].PutValue("Merged Header");
        cells["A2"].PutValue("Row 1");
        cells["B2"].PutValue("Data 1");
        cells["A3"].PutValue("Row 2");
        cells["B3"].PutValue("Data 2");

        // Merge cells A1:B1 to span across two columns
        // firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 2
        cells.Merge(0, 0, 1, 2);

        // Optionally, apply a style to the merged cell for better appearance
        Style style = cells["A1"].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        cells["A1"].SetStyle(style);

        // Save the workbook as PDF; merged cells will be preserved in the output
        workbook.Save("MergedCellsOutput.pdf", SaveFormat.Pdf);
    }
}

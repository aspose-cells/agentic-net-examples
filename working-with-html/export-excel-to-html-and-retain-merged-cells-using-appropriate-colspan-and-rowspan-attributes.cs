// Title: Export Excel to HTML with merged cells (colspan/rowspan) using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, merge a range (A2:C3), configure HtmlSaveOptions to output merged cells as proper colspan and rowspan attributes, and save the result as an HTML file while keeping the original layout intact.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | merged cells | colspan | rowspan | HtmlSaveOptions | MergeEmptyTdType | HtmlExportDataOptions | preserve layout | web report generation
// Common Searches: Aspose.Cells export merged cells to HTML | HTML export colspan rowspan Aspose .NET | keep merged cells when converting Excel to HTML | HtmlSaveOptions MergeEmptyTdType example | C# convert Excel workbook to HTML with merged ranges
// Developer Intent: Generate an HTML file from an Excel workbook that retains merged cell structures using Aspose.Cells.
// Use Cases: Create web‑ready reports from dynamically built workbooks with merged headers. | Display existing Excel documents on a website without losing complex cell merges. | Batch‑process multiple spreadsheets into HTML for documentation while preserving layout.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, ensuring merged cells use correct colspan and rowspan attributes. | Explain the impact of HtmlSaveOptions.MergeEmptyTdType and HtmlExportDataOptions on merged‑cell HTML output. | Show how to configure HtmlSaveOptions to preserve merged ranges when saving an Excel file as HTML in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, merge a range (A2:C3), configure HtmlSaveOptions to output merged cells as proper colspan and rowspan attributes, and save the result as an HTML file while keeping the original layout intact.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Merged Cell");

            // Merge a range of cells (A2:C3) – this will be represented with colspan/rowspan in HTML
            sheet.Cells.Merge(1, 0, 2, 3); // rows 1-2, columns 0-2 (A2:C3)

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Ensure merged cells are exported using default behavior (colspan/rowspan)
            htmlOptions.MergeEmptyTdType = MergeEmptyTdType.Default;

            // Export all data (default) – keeps merged cells intact
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

            // Save the workbook as HTML
            workbook.Save("ExportedWithMergedCells.html", htmlOptions);
        }
    }
}

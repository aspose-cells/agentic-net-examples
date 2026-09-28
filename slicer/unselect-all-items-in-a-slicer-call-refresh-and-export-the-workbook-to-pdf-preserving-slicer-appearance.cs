// Title: How to clear all slicer selections, refresh them, and export an Excel workbook to PDF with slicer graphics using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, iterates through every worksheet, sets ShowAllItems = true on each slicer, calls Refresh, and saves the workbook as a PDF with slicer visuals preserved. | Generate a method using Aspose.Cells that removes all slicer filters, updates the linked pivot tables, and creates a PDF output using PdfSaveOptions. | Provide a snippet that demonstrates unselecting slicer items across all sheets, invoking slicer.Refresh, and exporting the workbook to PDF while keeping the slicer shapes intact.
// Common Searches: Aspose.Cells C# clear slicer filter before converting Excel to PDF | How to refresh slicer‑linked pivot tables and keep slicer appearance in PDF export | Export workbook to PDF with slicer graphics using Aspose.Cells .NET
// Tags: reset slicer selections Aspose.Cells | slicer.Refresh update pivot table | PdfSaveOptions SheetSet.Visible | enumerate worksheet slicer collection C# | preserve slicer rendering during PDF conversion

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

// // Load the workbook, loop through each worksheet's slicer collection, set ShowAllItems = true to clear selections, call Refresh on each slicer, then save the workbook as a PDF using PdfSaveOptions (SheetSet.Visible) so slicer graphics are retained.
class SlicerPdfExport
{
    static void Main()
    {
        // Load an existing workbook that contains a slicer
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Process each slicer
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Unselect all items by showing all items (clears any filter)
                slicer.ShowAllItems = true;

                // Refresh the slicer (also refreshes the underlying pivot table)
                slicer.Refresh();
            }
        }

        // Prepare PDF save options – export all visible sheets (default) and keep slicer appearance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure the slicer is rendered as part of the sheet
            // No additional settings are required; slicers are included by default
            SheetSet = SheetSet.Visible
        };

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}

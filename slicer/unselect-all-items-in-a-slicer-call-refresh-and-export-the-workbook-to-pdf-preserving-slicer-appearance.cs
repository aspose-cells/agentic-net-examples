using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

class SlicerUnselectAndExport
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

                // Unselect all items by showing all items (no filter applied)
                slicer.ShowAllItems = true;

                // Refresh the slicer (also refreshes the underlying PivotTable)
                slicer.Refresh();
            }
        }

        // Prepare PDF save options (default options preserve slicer appearance)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF, preserving slicer visual appearance
        workbook.Save("output.pdf", pdfOptions);
    }
}
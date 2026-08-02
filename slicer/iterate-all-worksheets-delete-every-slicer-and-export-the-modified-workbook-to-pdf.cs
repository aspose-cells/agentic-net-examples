using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

class DeleteSlicersAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Remove all slicers from the collection
            slicers.Clear();
        }

        // Set PDF save options (optional customization can be added here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the modified workbook to a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
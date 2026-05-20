using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and remove all slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear the slicer collection for the current worksheet
                sheet.Slicers.Clear();
            }

            // Prepare PDF save options to include all sheets
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                SheetSet = SheetSet.All   // Export all sheets to the PDF
            };

            // Save the workbook as a consolidated PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
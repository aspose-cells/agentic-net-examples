using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace the path with your file)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = worksheet.Slicers;

                // Remove slicers starting from the last index to avoid shifting issues
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    slicers.RemoveAt(i);
                }
            }

            // Prepare PDF save options (default options are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the modified workbook as a PDF file
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
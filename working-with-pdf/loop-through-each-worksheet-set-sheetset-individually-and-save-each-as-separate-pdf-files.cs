using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – loop worksheets, save each as separate PDF
    class SheetSetLoopPdfExport
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Create PDF save options for the current sheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Set SheetSet to include only the current worksheet (0‑based index)
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Build output file name using the worksheet name
                string sheetName = workbook.Worksheets[i].Name;
                string outputPath = $"Pdf_{sheetName}.pdf";

                // Save the workbook as PDF with the specified SheetSet
                workbook.Save(outputPath, pdfOptions);
            }
        }
    }
}
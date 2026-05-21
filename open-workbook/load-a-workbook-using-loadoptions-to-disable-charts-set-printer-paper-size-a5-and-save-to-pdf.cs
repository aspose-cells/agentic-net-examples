using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    class LoadDisableChartsAndSavePdf
    {
        static void Main()
        {
            // Create LoadOptions and configure it
            LoadOptions loadOptions = new LoadOptions();

            // Disable loading of charts by ignoring useless shapes (charts are treated as shapes)
            loadOptions.IgnoreUselessShapes = true;

            // Set the default printer paper size to A5
            loadOptions.SetPaperSize(PaperSizeType.PaperA5);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Prepare PDF save options (default options are sufficient for this task)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
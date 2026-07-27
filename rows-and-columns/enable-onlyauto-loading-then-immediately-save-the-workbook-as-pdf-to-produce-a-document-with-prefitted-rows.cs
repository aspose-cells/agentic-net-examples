using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Initialize load options.
        LoadOptions loadOptions = new LoadOptions();

        // The OnlyAuto loading option is not available in the current Aspose.Cells API.
        // If a future version provides a property such as loadOptions.OnlyAuto,
        // it can be enabled here:
        // loadOptions.OnlyAuto = true;

        // Load the workbook with the specified options.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Configure PDF save options. Default behavior preserves row heights and auto‑fit.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure each sheet is rendered on its natural page layout.
            OnePagePerSheet = false
        };

        // Save the workbook as PDF.
        workbook.Save("output.pdf", pdfOptions);
    }
}
// Author: Example demonstrating loading a workbook and saving to PDF with Aspose.Cells.
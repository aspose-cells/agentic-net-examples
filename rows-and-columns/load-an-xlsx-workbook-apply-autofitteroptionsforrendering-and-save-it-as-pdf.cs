using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create AutoFitterOptions and enable rendering-specific auto‑fit
        AutoFitterOptions autoFitOptions = new AutoFitterOptions();
        autoFitOptions.ForRendering = true;

        // Assign the AutoFitterOptions to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = autoFitOptions;

        // Load the XLSX workbook using the configured LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Create default PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
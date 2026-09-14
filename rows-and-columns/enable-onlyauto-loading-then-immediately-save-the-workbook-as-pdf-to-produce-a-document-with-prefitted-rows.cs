// Title: Enable OnlyAuto row auto‑fitting on workbook load and export to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file using LoadOptions with AutoFitterOptions.OnlyAuto = true, then save the workbook as a PDF. | Configure AutoFitterOptions to auto‑fit only rows without custom height during workbook loading and generate a PDF output in C#. | Show how to apply OnlyAuto auto‑fit on load and directly export the workbook to PDF without extra formatting steps.
// Common Searches: Aspose.Cells C# load Excel with OnlyAuto auto‑fit rows and convert to PDF | How to use LoadOptions.AutoFitterOptions.OnlyAuto for PDF export in .NET | Convert Excel to PDF while preserving only automatically fitted row heights using Aspose.Cells | C# example of loading workbook with OnlyAuto row fitting then saving as PDF
// Tags: OnlyAuto auto‑fit rows load option | Aspose.Cells PDF export with auto‑fitted rows | C# load workbook with AutoFitterOptions | Excel to PDF conversion using OnlyAuto fitting | AutoFitterOptions row height auto‑fit

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsOnlyAutoLoadToPdf
{
    // Loads an Excel file with LoadOptions.AutoFitterOptions.OnlyAuto set to true so only rows without custom heights are auto‑fitted, then saves the workbook directly to a PDF file using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path for the resulting PDF file
            string outputPath = "output.pdf";

            // ------------------------------------------------------------
            // 1. Configure load options to enable OnlyAuto auto‑fitting.
            // ------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFitterOptions = new AutoFitterOptions();
            // Only rows whose height is not customed will be auto‑fitted.
            loadOptions.AutoFitterOptions.OnlyAuto = true;

            // ------------------------------------------------------------
            // 2. Load the workbook with the configured options.
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // ------------------------------------------------------------
            // 3. Prepare PDF save options (default options are sufficient).
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // ------------------------------------------------------------
            // 4. Save the workbook directly to PDF.
            //    The rows have already been auto‑fitted during loading.
            // ------------------------------------------------------------
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook loaded with OnlyAuto fitting and saved as PDF to '{outputPath}'.");
        }
    }
}

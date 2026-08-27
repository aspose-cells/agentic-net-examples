// Title: Load an XLSX workbook with AutoFitterOptions.ForRendering and convert it to PDF using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with AutoFitterOptions.ForRendering enabled, then save the workbook as a PDF using Aspose.Cells. | Apply rendering‑mode auto‑fit during workbook loading and export the result with default PdfSaveOptions in C#.
// Common Searches: Aspose.Cells C# load Excel file with AutoFitterOptions rendering mode and export to PDF | How to enable AutoFitterOptions.ForRendering when converting XLSX to PDF in .NET | C# convert XLSX to PDF with column autofit for rendering using Aspose.Cells
// Tags: AutoFitterOptions.ForRendering load option | XLSX to PDF conversion Aspose.Cells | rendering mode column autofit Aspose.Cells | PdfSaveOptions default usage | Workbook.Save PDF Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitPdfDemo
{
    // The program loads 'input.xlsx' using LoadOptions that contain AutoFitterOptions with ForRendering set to true, then saves the workbook as 'output.pdf' using default PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create AutoFitterOptions and enable rendering mode
            AutoFitterOptions autoFitOptions = new AutoFitterOptions();
            autoFitOptions.ForRendering = true; // Apply rendering‑specific auto‑fit

            // Assign the options to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFitterOptions = autoFitOptions;

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Create PDF save options (default constructor)
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the save options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfSaveOptions);

            Console.WriteLine($"Workbook converted to PDF successfully: {outputPath}");
        }
    }
}

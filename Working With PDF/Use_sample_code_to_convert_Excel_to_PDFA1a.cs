using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExcelToPdfA1a
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF/A‑1a file
        string destPath = "output_pdfa1a.pdf";

        // Load the Excel workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure PDF save options for PDF/A‑1a compliance
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA1a; // Use PDF/A‑1a standard

        // Save the workbook as PDF/A‑1a
        workbook.Save(destPath, saveOptions);

        Console.WriteLine("Excel file successfully converted to PDF/A‑1a.");
    }
}
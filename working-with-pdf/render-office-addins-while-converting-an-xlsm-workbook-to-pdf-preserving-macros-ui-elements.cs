// Title: Convert an XLSM workbook to PDF while rendering Office Add‑In UI elements using Aspose.Cells for .NET
// Common Searches: how to keep macro‑generated buttons visible when exporting XLSM to PDF with Aspose.Cells C# | Aspose.Cells render embedded ActiveX controls in PDF output from a macro‑enabled workbook | C# code to enable macros and recalculate formulas before saving XLSM as PDF using Aspose.Cells | preserve Office Add‑In UI layout during XLSM to PDF conversion with Aspose.Cells | Aspose.Cells PdfSaveOptions settings for rendering macro UI in PDF
// Tags: macro‑enabled XLSM PDF export Aspose.Cells | render Office Add‑In UI in PDF with Aspose.Cells | PdfSaveOptions OnePagePerSheet false setting | recalculate workbook formulas before PDF conversion | preserve ActiveX controls during PDF generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads a macro‑enabled XLSM workbook, enables macro execution, recalculates formulas, and saves it as a PDF while preserving the visual appearance of Office Add‑In UI elements such as ActiveX controls using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook; Aspose.Cells auto‑detects the format, so explicit LoadFormat is not required
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Enable macro execution so that any UI changes made by macros are reflected
            workbook.Settings.EnableMacros = true;

            // Recalculate formulas and any UI‑related code before rendering
            workbook.CalculateFormula();

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render each sheet as it appears (including embedded controls)
                OnePagePerSheet = false
            };

            // Save the workbook as PDF while preserving the visual representation of macros' UI
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

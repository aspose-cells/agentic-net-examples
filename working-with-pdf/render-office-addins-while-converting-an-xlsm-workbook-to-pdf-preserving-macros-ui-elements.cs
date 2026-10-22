// Title: Convert an XLSM workbook with Office Add‑Ins to PDF while preserving macro UI elements using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a macro‑enabled XLSM file with Aspose.Cells, sets EnableMacros = true, and saves it as a PDF using PdfSaveOptions that keep the original sheet layout. | Show how to configure PdfSaveOptions in Aspose.Cells to disable OnePagePerSheet and AllColumnsInOnePagePerSheet for XLSM‑to‑PDF conversion. | Write robust error handling that verifies the input XLSM file exists before conversion and logs success or failure messages.
// Common Searches: Aspose.Cells C# convert macro enabled XLSM to PDF preserving ActiveX controls | How to keep original sheet layout when saving XLSM as PDF with Aspose.Cells | Enable macros on workbook load Aspose.Cells .NET example | PdfSaveOptions OnePagePerSheet false usage in C# | Check file existence before XLSM to PDF conversion Aspose.Cells
// Tags: XLSM to PDF conversion with macro UI preservation | Aspose.Cells PdfSaveOptions layout configuration | render Office Add‑Ins as static graphics Aspose.Cells | enable macros on workbook load Aspose.Cells | C# file existence validation for Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the XLSM file exists, loads it with macros enabled, configures PdfSaveOptions to retain the original sheet layout (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false), and saves the workbook as a PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the macro‑enabled workbook (XLSM)
            Workbook workbook = new Workbook(inputPath);

            // Preserve macros during load (prevents their removal)
            workbook.Settings.EnableMacros = true;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep the original sheet layout (do not force one page per sheet)
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = false
                // Note: RenderActiveXControls and RenderFormControls are not available in this version;
                // the default behavior renders them as static graphics.
            };

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

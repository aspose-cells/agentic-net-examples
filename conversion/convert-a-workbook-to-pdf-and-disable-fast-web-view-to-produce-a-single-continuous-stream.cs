// Title: Convert Excel to PDF without Fast Web View using Aspose.Cells for .NET
// Description: Loads an XLSX workbook, configures PdfSaveOptions to disable Fast Web View (or EnableFastWebView) via reflection, and saves a single‑stream PDF. Includes file‑existence checks and version‑agnostic handling.
// Keywords: Aspose.Cells PDF conversion | disable Fast Web View | FastWebView false | EnableFastWebView false | single stream PDF | C# Aspose.Cells example | Excel to PDF Aspose | reflection PdfSaveOptions | continuous PDF stream
// Common Searches: Aspose.Cells disable fast web view when saving PDF | How to create a single‑stream PDF from Excel using Aspose | C# set FastWebView false Aspose.Cells | PdfSaveOptions FastWebView property reflection | Convert .xlsx to PDF without incremental pages
// Developer Intent: Generate a PDF from an Excel workbook while ensuring Fast Web View is turned off to produce a continuous, non‑incremental stream.
// Use Cases: Deliver PDFs for web download that must load as a single stream. | Create archival PDFs required for digital signing or compliance. | Support projects that run on older Aspose.Cells versions with differing property names.
// AI Prompts: Write C# code with Aspose.Cells that converts an .xlsx file to PDF and disables Fast Web View, handling both FastWebView and EnableFastWebView via reflection. | Explain why disabling Fast Web View results in a single‑stream PDF and when this is needed. | Provide error handling for missing input files and fallback logic if the Fast Web View property is unavailable.

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook, configures PdfSaveOptions to disable Fast Web View (or EnableFastWebView) via reflection, and saves a single‑stream PDF. Includes file‑existence checks and version‑agnostic handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Disable Fast Web View if the property is available in the referenced version
            // Some older versions use EnableFastWebView instead of FastWebView
            // Use reflection to set the appropriate property without breaking compilation
            var fastWebViewProp = typeof(PdfSaveOptions).GetProperty("FastWebView");
            if (fastWebViewProp != null && fastWebViewProp.CanWrite)
            {
                fastWebViewProp.SetValue(pdfOptions, false);
            }
            else
            {
                var enableFastWebViewProp = typeof(PdfSaveOptions).GetProperty("EnableFastWebView");
                if (enableFastWebViewProp != null && enableFastWebViewProp.CanWrite)
                {
                    enableFastWebViewProp.SetValue(pdfOptions, false);
                }
            }

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

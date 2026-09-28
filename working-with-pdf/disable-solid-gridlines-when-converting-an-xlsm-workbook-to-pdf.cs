// Title: How to hide worksheet gridlines when converting an XLSM workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a macro‑enabled XLSM workbook with Aspose.Cells, disables gridlines on every worksheet, and saves it as a PDF. | Show how to configure PdfSaveOptions in Aspose.Cells so the generated PDF does not display any gridlines. | Explain the steps to iterate through all worksheets in a Workbook object and set IsGridlinesVisible = false before exporting to PDF.
// Common Searches: Aspose.Cells C# hide gridlines before saving XLSM as PDF | disable solid gridlines in PDF output from macro enabled workbook using Aspose.Cells | C# code to turn off worksheet gridlines when exporting to PDF with Aspose.Cells | PdfSaveOptions gridlines visibility Aspose.Cells .NET
// Tags: Aspose.Cells hide worksheet gridlines PDF export | C# set IsGridlinesVisible false Aspose.Cells | PdfSaveOptions gridlines control Aspose.Cells | export XLSM to PDF without gridlines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSM workbook, disables gridlines on each worksheet by setting IsGridlinesVisible to false, and saves the workbook as a PDF using default PdfSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Disable gridlines for each worksheet (affects PDF rendering)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsGridlinesVisible = false;
            }

            // Configure PDF save options (default options are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

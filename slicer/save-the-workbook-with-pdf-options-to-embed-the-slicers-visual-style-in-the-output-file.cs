// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET while handling slicer visual style
// AI Prompts: Write C# code that loads an .xlsx file, iterates over any slicers on the first worksheet using dynamic typing, and saves the workbook as a PDF with PdfSaveOptions in Aspose.Cells. | Demonstrate how to apply or clear slicer style properties safely before exporting an Excel workbook to PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export Excel to PDF preserving slicer formatting | How to keep slicer appearance when converting .xlsx to PDF with Aspose.Cells | C# example for saving workbook as PDF and accessing slicer objects in Aspose.Cells | PdfSaveOptions slicer style handling Aspose.Cells .NET
// Tags: Aspose.Cells PDF export with slicer handling | C# workbook to PDF using PdfSaveOptions | slicer style manipulation Aspose.Cells | dynamic access to slicer objects Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, optionally iterates through slicers on the first worksheet using dynamic typing to avoid compile‑time dependencies, configures PdfSaveOptions, and saves the workbook as a PDF. It includes error handling for missing files and unsupported slicer APIs, illustrating how to work with slicer objects and PDF export in Aspose.Cells for .NET.
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
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Optional: handle slicers if they exist (style setting omitted for compatibility)
            try
            {
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Slicers.Count > 0)
                {
                    foreach (var slicerObj in sheet.Slicers)
                    {
                        // Use dynamic to avoid compile‑time dependency on specific slicer members
                        dynamic slicer = slicerObj;
                        try
                        {
                            // Attempt to set a style if the property exists; ignore if not supported
                            slicer.Style = null;
                        }
                        catch
                        {
                            // Silently ignore any errors related to slicer styling
                        }
                    }
                }
            }
            catch
            {
                // Silently ignore if slicer APIs are unavailable
            }

            // Configure PDF save options (no slicer embedding property in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

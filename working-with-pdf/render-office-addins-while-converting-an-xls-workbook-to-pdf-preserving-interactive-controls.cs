// Title: Convert a legacy XLS workbook to PDF with Aspose.Cells for .NET while attempting to retain Office Add‑In controls (form fields and OLE objects)
// AI Prompts: Write C# code that loads an .xls file with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet to false, and saves the workbook as a PDF, including comments about the unavailable PreserveFormFields and PreserveOleObjects options. | Explain how to implement robust file‑existence verification and exception handling around an Aspose.Cells XLS‑to‑PDF conversion in C#. | Describe how to detect that certain PDF save options (e.g., PreserveFormFields) are not supported in the current Aspose.Cells API version and suggest alternative approaches.
// Common Searches: Aspose.Cells .NET convert .xls to PDF without losing form fields | C# PdfSaveOptions OnePagePerSheet false example | Why is PreserveFormFields missing in Aspose.Cells PDF conversion | Render Office Add‑In controls when saving Excel as PDF using Aspose.Cells | Handle file not found error during Aspose.Cells workbook conversion
// Tags: Aspose.Cells XLS to PDF conversion | PdfSaveOptions OnePagePerSheet setting | preserve interactive controls Aspose.Cells | C# file existence check Aspose.Cells | exception handling Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies that input.xls exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with OnePagePerSheet set to false, notes that PreserveFormFields and PreserveOleObjects are not available, and saves the workbook as output.pdf while providing comprehensive error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.pdf";

        // Verify that the source workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each worksheet on a separate PDF page if desired
                OnePagePerSheet = false
                // Note: PreserveFormFields and PreserveOleObjects are not available in the current API version
            };

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

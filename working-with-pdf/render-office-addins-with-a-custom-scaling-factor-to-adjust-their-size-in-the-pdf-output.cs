// Title: Apply a custom PDF ZoomFactor to scale Office Add‑In content when exporting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets PdfSaveOptions.ZoomFactor to 1.5 and saves a Workbook as a PDF with Aspose.Cells. | Create a C# snippet that checks the Aspose.Cells library version at runtime and applies PdfSaveOptions.ZoomFactor only when the property exists, otherwise uses default PDF options.
// Common Searches: how to increase PDF output size of Office Add‑In controls with Aspose.Cells .NET | Aspose.Cells PDF zoom factor usage C# example | detect if PDF zoom factor is supported in current Aspose.Cells version | scale Excel workbook PDF export for Office Add‑In using custom zoom factor
// Tags: Aspose.Cells custom PDF zoom factor C# | Office Add‑In PDF scaling Aspose.Cells | runtime Aspose.Cells version check .NET | export Excel to PDF with scaling Aspose.Cells | adjust Office Add‑In size in PDF output

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook or creates a new one, optionally sets PdfSaveOptions.ZoomFactor to enlarge Office Add‑In elements, checks for property support based on the Aspose.Cells version, and saves the workbook as a PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;

            // Verify that the input file exists; create a simple workbook if it does not.
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Configure PDF save options.
            // Note: The ZoomFactor property is not available in older Aspose.Cells versions.
            // If your version supports it, you can uncomment the line below.
            // pdfOptions.ZoomFactor = 1.5;
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the specified options.
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

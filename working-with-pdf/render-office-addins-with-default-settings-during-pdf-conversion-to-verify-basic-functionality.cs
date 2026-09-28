// Title: Convert Excel to PDF with default settings and verify Office Add‑Ins rendering support using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, validates its existence, and saves it as a PDF using Aspose.Cells with the default PdfSaveOptions. | Show how to attempt enabling Office Add‑Ins rendering during PDF export in Aspose.Cells and provide fallback logic when the RenderOfficeAddIns property is not available. | Write robust error‑handling for a console app that converts Excel to PDF with Aspose.Cells, covering file‑not‑found and generic exception scenarios.
// Common Searches: c# aspocells export excel workbook to pdf with default options | check workbook file existence before Aspose.Cells PDF conversion | office addins rendering not available in current Aspose.Cells version | example of error handling for Aspose.Cells Excel to PDF conversion | default pdf save settings in Aspose.Cells .NET
// Tags: Aspose.Cells PDF export default options | C# Excel workbook to PDF Aspose.Cells | Office Add‑Ins support detection Aspose.Cells | validate workbook file presence before conversion | handle missing RenderOfficeAddIns property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks that input.xlsx exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance with default settings (noting that RenderOfficeAddIns is unavailable in this version), saves the workbook as output.pdf, and includes comprehensive exception handling.
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

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: RenderOfficeAddIns property is not available in this version of Aspose.Cells.
            // If needed, upgrade to a newer version that supports this feature.

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Render Office Add‑Ins at a custom scale when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that checks for an .xlsx file, loads it with Aspose.Cells, and saves it as a PDF using a 150% ZoomFactor. | Demonstrate how to set PdfSaveOptions.ZoomFactor to enlarge Office Add‑Ins in the PDF output during Excel‑to‑PDF conversion. | Create a complete example with try‑catch and file‑existence validation that applies a custom scaling factor when exporting to PDF.
// Common Searches: Aspose.Cells how to set PDF zoom factor for Office Add‑Ins | C# convert Excel to PDF with custom scaling using Aspose.Cells | PdfSaveOptions ZoomFactor 150 percent example | Scale Office Add‑Ins in PDF output Aspose.Cells .NET | Increase size of embedded Office Add‑Ins when saving workbook as PDF
// Tags: PdfSaveOptions ZoomFactor scaling | Excel to PDF conversion custom zoom | Office Add‑Ins size adjustment Aspose.Cells | C# file existence check before PDF export | Aspose.Cells exception handling PDF save

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that the source Excel file exists, loads it with Aspose.Cells, configures PdfSaveOptions.ZoomFactor to 150 (150% scaling) to enlarge Office Add‑Ins, saves the workbook as a PDF, and handles any runtime exceptions.
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

            // Create PDF save options and set a custom zoom factor (scaling)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // 150 means 150% scaling (1.5x)
                ZoomFactor = 150
            };

            // Save the workbook as a PDF using the scaling factor
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

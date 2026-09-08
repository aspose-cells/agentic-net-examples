// Title: Create a PDF/A‑1b compliant workbook from Aspose.Cells in C# using PdfSaveOptions
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, adds sample data, and saves it as a PDF/A‑1b document by configuring PdfSaveOptions. | Show how to set PdfSaveOptions.Compliance to PdfA1b and adjust additional PDF export settings before calling Workbook.Save in a .NET application.
// Common Searches: how to save an Aspose.Cells workbook as PDF/A-1b using C# | Aspose.Cells PdfSaveOptions compliance setting example | C# export Excel to PDF with PDF/A-1b compliance Aspose.Cells | set PDF/A-1b compliance when converting workbook to PDF in Aspose.Cells | configure PDF export options for Aspose.Cells workbook in .NET
// Tags: Aspose.Cells PDF/A-1b export C# | PdfSaveOptions compliance configuration | save workbook as PDF with Aspose.Cells | PDF/A-1b conversion Aspose.Cells | C# Aspose.Cells PDF export settings

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Aspose.Cells workbook, inserts a value into cell A1, configures PdfSaveOptions with PDF/A‑1b compliance, and saves the workbook as a PDF file while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Prepare PDF save options with PDF/A-1b compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

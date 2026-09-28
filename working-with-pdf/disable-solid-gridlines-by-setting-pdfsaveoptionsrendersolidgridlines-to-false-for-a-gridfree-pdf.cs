// Title: Generate a grid‑free PDF from an Aspose.Cells workbook by hiding worksheet gridlines in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, sets worksheet.IsGridlinesVisible to false, and saves it as a PDF using PdfSaveOptions. | Show how to export an Excel workbook to a PDF without any gridlines in Aspose.Cells for .NET, including workbook setup and PDF save options.
// Common Searches: Aspose.Cells hide worksheet gridlines before PDF export C# | how to generate a PDF from Excel without showing gridlines in Aspose.Cells C# | PdfSaveOptions save workbook as PDF without gridlines example
// Tags: Aspose.Cells hide worksheet gridlines | C# export Excel to PDF without gridlines | PdfSaveOptions gridline‑free PDF | grid‑free PDF generation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, adds sample data, disables worksheet gridlines via IsGridlinesVisible = false, configures PdfSaveOptions, and saves the workbook as a PDF file that contains no gridlines.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data to illustrate the result
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(12345);

            // Hide gridlines in the worksheet before saving to PDF
            sheet.IsGridlinesVisible = false;

            // PDF save options (no gridline-specific property needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            string outputPath = "GridFreeOutput.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

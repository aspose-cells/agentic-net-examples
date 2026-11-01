// Title: Generate a PDF/A‑1b compliant document from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that builds an Excel workbook, fills cells with sample data, and saves it as a PDF/A‑1b file using Aspose.Cells PdfSaveOptions. | Demonstrate how to configure Aspose.Cells PdfSaveOptions to enforce PDF/A‑1b compliance when exporting a workbook to PDF in a .NET application.
// Common Searches: how to export Excel to PDF/A 1b using Aspose.Cells in C# | set PDF/A compliance level with Aspose.Cells PdfSaveOptions .NET | sample code for creating PDF/A‑1b from workbook Aspose | archival PDF generation from Excel with Aspose.Cells example
// Tags: aspocells pdfa1b export c# | pdfsaveoptions compliance setting aspocells | excel workbook to archival pdf using aspocells | c# generate pdfa compliant document from worksheet | save workbook as pdfa1b with aspocells rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, populates it with sample product data, configures PdfSaveOptions with PdfCompliance.PdfA1b for PDF/A‑1b compliance, ensures the output directory exists, and saves the workbook as an archival PDF/A‑1b file named SamplePdfA.pdf.
class PdfAConversionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the workbook with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // Configure PDF save options for PDF/A compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the desired PDF/A compliance level (PDF/A-1b in this case)
                Compliance = PdfCompliance.PdfA1b
                // EmbedStandardFonts and Title are not supported in this version of PdfSaveOptions
            };

            // Define output file path
            string outputPath = "SamplePdfA.pdf";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF/A compliant document
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF/A file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

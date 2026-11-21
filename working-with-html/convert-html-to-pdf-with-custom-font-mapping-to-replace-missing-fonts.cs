// Title: Convert HTML to PDF with Aspose.Cells in C# – validate input file and create output folder
// Common Searches: C# Aspose.Cells example for converting a local HTML file to PDF with error handling | How to use HtmlLoadOptions and PdfSaveOptions in Aspose.Cells .NET | Load HTML into an Aspose.Cells workbook and export as PDF in C# | Create missing output directory before saving PDF with Aspose.Cells | Add custom FontSettings to replace missing fonts when converting HTML to PDF using Aspose.Cells
// Tags: Aspose.Cells HtmlLoadOptions PDF export | C# workbook save as PDF Aspose.Cells | HTML to PDF conversion .NET Aspose.Cells | file existence check before Aspose.Cells conversion | create output directory for PDF generation | custom FontSettings missing font mapping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample validates that the source HTML file exists, creates the destination folder if needed, loads the HTML into an Aspose.Cells Workbook using HtmlLoadOptions, and saves it as a PDF with PdfSaveOptions. It includes exception handling and notes that FontSettings can be added to map missing fonts during conversion.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Paths
        string htmlPath = @"C:\Input\sample.html";
        string pdfPath = @"C:\Output\sample.pdf";

        try
        {
            // Verify input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load HTML (no custom FontSettings used here)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save as PDF
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            workbook.Save(pdfPath, saveOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

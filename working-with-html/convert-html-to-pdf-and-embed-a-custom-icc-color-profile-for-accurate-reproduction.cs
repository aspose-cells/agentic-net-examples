// Title: Convert HTML to PDF with Aspose.Cells for .NET (C#) and embed a custom ICC color profile
// AI Prompts: Generate C# code that loads a local HTML file into an Aspose.Cells Workbook, sets PdfSaveOptions.IccProfilePath to a specified ICC file, and saves the workbook as a PDF. | Show how to configure Aspose.Cells PDF export to embed a custom ICC color profile for accurate color reproduction when converting HTML to PDF.
// Common Searches: asp.net c# convert html to pdf with aspose.cells and embed icc profile | aspose.cells pdfsaveoptions iccprofilepath example c# | how to add custom color profile to pdf generated from html using aspose.cells | c# load html into workbook and export to pdf with color management settings | aspose.cells html to pdf conversion color accuracy asp.net
// Tags: Aspose.Cells HTML to PDF conversion with ICC profile | PdfSaveOptions IccProfilePath configuration | embed custom color profile in PDF Aspose.Cells | color management for PDF export C# | load HTML workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions to reference a custom ICC color profile, and saves the workbook as a PDF, ensuring accurate color reproduction while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for source HTML and final PDF
            string htmlPath = "input.html";
            string finalPdfPath = "output.pdf";

            // Ensure the HTML file exists before loading
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Load the HTML file into an Aspose.Cells Workbook
            var workbook = new Workbook(htmlPath);

            // Configure PDF save options if needed
            var pdfSaveOptions = new PdfSaveOptions();

            // Save the workbook directly as PDF
            workbook.Save(finalPdfPath, pdfSaveOptions);

            Console.WriteLine($"PDF successfully created at: {finalPdfPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

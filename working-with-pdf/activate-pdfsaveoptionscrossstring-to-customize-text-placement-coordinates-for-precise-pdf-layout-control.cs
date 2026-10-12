// Title: How to enable PdfSaveOptions.CrossString for precise text placement when exporting an Aspose.Cells workbook to PDF in C#
// AI Prompts: Write C# code that sets the PdfSaveOptions.CrossString property with X and Y values to control where cell text appears in the generated PDF. | Show how to combine PdfSaveOptions.OnePagePerSheet with the CrossString property to produce a multi‑sheet PDF with exact text positioning. | Explain how enabling CrossString affects PDF rendering of cell values and how it can be used to fine‑tune layout in Aspose.Cells.
// Common Searches: Aspose.Cells C# set PdfSaveOptions.CrossString to adjust text coordinates in PDF export | how to control cell text placement when saving a workbook as PDF using Aspose.Cells | custom PDF layout for multi‑sheet workbooks with Aspose.Cells .NET | example of using CrossString for precise PDF text positioning in C#
// Tags: PdfSaveOptions.CrossString coordinate control | Aspose.Cells PDF text positioning | C# export workbook to PDF with custom layout | OnePagePerSheet PDF generation Aspose.Cells | custom cell text placement Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds data, configures PdfSaveOptions with OnePagePerSheet and activates the CrossString property to specify X/Y coordinates for text placement, ensures the output directory exists, and saves the workbook as a precisely laid‑out PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B2"].PutValue("World");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render each worksheet on a separate PDF page
                OnePagePerSheet = true
            };

            // Define output file path
            string outputPath = "output.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

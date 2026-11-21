// Title: Convert an Excel workbook with WordArt shapes to PDF while preserving gradient fills using Aspose.Cells for .NET
// Common Searches: c# aspnet convert excel file containing wordart to pdf preserving gradient fill | aspocells export workbook with wordart shapes to pdf with accurate colors | how to keep wordart gradient shading when saving excel as pdf in .net | pdfsaveoptions settings for rendering wordart correctly in aspocells | convert xlsx to pdf with shape formatting using aspocells c#
// Tags: Aspose.Cells PDFSaveOptions shape rendering | Excel WordArt to PDF conversion | gradient fill preservation in PDF export | C# workbook to PDF with shapes | Aspose.Cells preserve WordArt formatting

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook that may contain WordArt objects, configures PdfSaveOptions, and saves the file as a PDF, ensuring that gradient fills and other shape formatting are retained in the output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (may contain WordArt objects)
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: RasterizationResolution property is not available in the current Aspose.Cells version.
            // If needed, adjust other options such as OnePagePerSheet, etc.

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

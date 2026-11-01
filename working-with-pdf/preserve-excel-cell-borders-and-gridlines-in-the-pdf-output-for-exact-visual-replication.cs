// Title: How to retain Excel cell borders and gridlines when converting to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# preserve cell borders in PDF export | Save Excel workbook to PDF with gridlines using Aspose.Cells .NET | C# convert .xlsx to PDF while keeping borders and gridlines Aspose | PdfSaveOptions default settings keep Excel gridlines Aspose.Cells | Ensure exact visual layout when exporting Excel to PDF in C#
// Tags: Aspose.Cells PDF export preserve borders | PdfSaveOptions retain gridlines | Excel to PDF visual fidelity C# | Workbook.Save PDF options Aspose.Cells | C# Excel PDF conversion exact layout

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object (which by default keeps gridlines and cell borders), and saves the workbook as a PDF, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default options preserve gridlines and cell borders)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

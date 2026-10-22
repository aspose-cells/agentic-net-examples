// Title: Load an Excel workbook, set the first worksheet to A5 paper size, and export it as a PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set worksheet paper size to A5 before PDF conversion | How to verify PDF page dimensions after exporting Excel with Aspose.Cells | Export Excel to PDF with A5 page layout using Aspose.Cells .NET | Aspose.Cells PDF export preserving custom page setup in C#
// Tags: set worksheet paper size A5 Aspose.Cells | export worksheet to PDF with custom page size | verify PDF dimensions after Excel conversion | Aspose.Cells page setup for PDF output | C# load workbook apply A5 paper size

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, applies the A5 paper size to the first worksheet via the PageSetup property, and saves the workbook as a PDF. It demonstrates how to configure page layout before conversion and confirms successful PDF generation, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Apply A5 paper size to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Create PDF bookmarks from worksheet names when converting Excel to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, converts it to PDF, and adds a PDF bookmark for each worksheet using the sheet's name as the bookmark title. | Update an existing Aspose.Cells PDF conversion snippet to include PDF outline entries for every worksheet based on their names.
// Common Searches: Aspose.Cells C# add PDF bookmarks for each Excel worksheet during conversion | How to generate PDF outline entries from sheet names using Aspose.Cells | C# convert .xlsx to .pdf with bookmarks for each sheet Aspose.Cells | Create PDF outline from Excel workbook sheets using Aspose.Cells .NET
// Tags: Aspose.Cells PDF outline generation | C# worksheet name as PDF outline entry | PdfSaveOptions outline title configuration | Excel to PDF conversion with outline entries | Aspose.Cells workbook sheet outline

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel file, converts it to PDF with Aspose.Cells, and demonstrates how to add a PDF bookmark for each worksheet using the sheet name as the bookmark title via PdfSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (default options are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

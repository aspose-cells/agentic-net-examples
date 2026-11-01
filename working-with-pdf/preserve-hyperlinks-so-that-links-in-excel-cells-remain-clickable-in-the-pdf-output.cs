// Title: Generate a PDF from an Excel workbook in C# with Aspose.Cells while keeping cell hyperlinks clickable
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to keep hyperlinks active, and saves it as a PDF. | Show how to verify an Excel file exists before converting it to PDF using Aspose.Cells, ensuring hyperlinks remain functional in the output. | Demonstrate setting up Aspose.Cells PdfSaveOptions in a .NET console app to preserve clickable URLs when exporting a workbook to PDF.
// Common Searches: asp.net convert excel to pdf retain hyperlinks Aspose.Cells | c# Aspose.Cells PdfSaveOptions keep hyperlink active in PDF output | how to export workbook to PDF with clickable links using Aspose.Cells .NET | preserve Excel cell URLs when saving as PDF with Aspose.Cells library
// Tags: Aspose.Cells PDF conversion hyperlink retention | PdfSaveOptions preserve active links | C# Excel to PDF export preserving URLs | Aspose.Cells workbook save as PDF clickable hyperlinks | Excel hyperlink preservation in PDF using .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example checks that an input .xlsx file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (default settings keep hyperlinks active), and saves the workbook as a PDF, producing a document where all cell hyperlinks remain clickable.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings keep hyperlinks active)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

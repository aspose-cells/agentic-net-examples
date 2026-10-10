// Title: How to keep Excel cell hyperlinks clickable when converting a workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Show a C# example that converts an Excel workbook to PDF using Aspose.Cells while ensuring all cell hyperlinks remain clickable. | Provide code to add a hyperlink to a worksheet cell and verify that the link is preserved after saving the workbook as a PDF with PdfSaveOptions. | Generate a snippet that loads an .xlsx file, inserts a clickable URL, and exports to PDF with hyperlink support in Aspose.Cells.
// Common Searches: C# Aspose.Cells keep hyperlinks when saving workbook as PDF | preserve clickable links in PDF generated from Excel using Aspose.Cells .NET | how to add and retain Excel cell hyperlink after PDF conversion with Aspose.Cells | PdfSaveOptions hyperlink preservation example in C#
// Tags: Aspose.Cells preserve hyperlinks PDF conversion | C# PdfSaveOptions hyperlink support | add hyperlink to Excel cell Aspose.Cells | export workbook to PDF with active links | hyperlink retention during Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing Excel file, optionally adds a hyperlink to cell A1, and saves the workbook as a PDF using Aspose.Cells. The PdfSaveOptions preserve the hyperlink, so the generated PDF contains clickable URLs.
class HyperlinkPreservationExample
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // OPTIONAL: Add a hyperlink to demonstrate preservation
            // This adds a hyperlink in cell A1 pointing to https://example.com
            Worksheet sheet = workbook.Worksheets[0];

            // Add hyperlink to cell A1 (row 0, column 0) covering 1 row and 1 column
            int hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com");
            Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = "Example Site";

            // Configure PDF save options (hyperlinks are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

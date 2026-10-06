// Title: Render Excel slicer as a static image in PDF by clearing slicers with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, removes all slicers from the first worksheet, and saves the result as a PDF so the slicer is rendered as a non‑interactive image. | Explain the steps to disable slicer interactivity when exporting an Excel sheet to PDF using Aspose.Cells in a .NET project. | Show how to verify the input file, clear slicers, and invoke Workbook.Save with SaveFormat.Pdf to produce a static PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF without slicer controls | how to render Excel slicer as image in PDF using .NET | remove slicers before saving workbook as PDF with Aspose.Cells | static PDF conversion of Excel worksheet containing slicers in C# | convert Excel slicer to non‑interactive image in PDF Aspose.Cells
// Tags: Aspose.Cells clear slicers for PDF export | C# static PDF generation from Excel worksheet | remove worksheet slicers Aspose.Cells | export Excel slicer as image PDF | Aspose.Cells SaveFormat.Pdf without interactivity

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing XLSX file with Aspose.Cells, checks for slicers on the first worksheet, clears them to prevent interactivity, and then saves the workbook as a static PDF where the slicer appears as a regular image.
class SlicerToStaticPdf
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // If the worksheet contains slicers, remove them to make the PDF static
            if (sheet.Slicers.Count > 0)
            {
                // Remove all slicers (or just the first one if preferred)
                sheet.Slicers.Clear();
            }

            // Save the workbook as a static PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

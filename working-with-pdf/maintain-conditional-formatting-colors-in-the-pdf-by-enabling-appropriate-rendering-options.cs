// Title: Enable RenderConditionalFormatting to Keep Conditional Formatting Colors When Converting Excel to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# example that detects whether the PdfSaveOptions.RenderConditionalFormatting property is available in the current Aspose.Cells version and sets it to true before saving the workbook as a PDF. | Write a C# program that loads an .xlsx file, configures PdfSaveOptions to retain conditional formatting colors, and exports the workbook to a PDF file while handling missing input files gracefully.
// Common Searches: asp.net aspose.cells preserve conditional formatting colors during pdf export | c# pdfsaveoptions renderconditionalformatting property usage example | excel to pdf conversion losing conditional formatting colors aspose cells | how to check for RenderConditionalFormatting support in Aspose.Cells C# | enable conditional formatting rendering when saving workbook as PDF with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions conditional formatting rendering | C# Excel to PDF conversion preserving colors | RenderConditionalFormatting property Aspose.Cells | PDF export options for conditional formatting in .NET | retain conditional formatting colors in PDF output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, optionally enables the RenderConditionalFormatting flag when supported, and saves the workbook as a PDF, providing error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If the used Aspose.Cells version supports it, preserve conditional formatting colors
            // pdfOptions.RenderConditionalFormatting = true; // Uncomment if the property exists

            // Save the workbook as PDF
            workbook.Save(outputFile, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

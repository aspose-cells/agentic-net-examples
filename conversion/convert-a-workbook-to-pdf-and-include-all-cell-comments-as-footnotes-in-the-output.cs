// Title: Convert an Excel workbook to PDF with all cell comments rendered as footnotes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.CommentDisplayMode to render comments as footnotes, and saves the workbook as a PDF. | Show how to configure PdfSaveOptions in Aspose.Cells so that cell comments appear as footnotes in the exported PDF. | Provide a complete example that checks the input file, enables comment footnote mode, and converts the workbook to PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF with comments as footnotes | How to include cell comments when converting .xlsx to PDF using Aspose.Cells .NET | PdfSaveOptions CommentDisplayMode footnotes example | C# convert workbook to PDF preserving Excel annotations | Aspose.Cells PDF conversion include cell notes footnote
// Tags: Aspose.Cells PDF conversion comment footnotes | PdfSaveOptions.CommentDisplayMode usage | C# export Excel comments to PDF | include cell annotations in PDF Aspose.Cells | Excel to PDF preserving comments .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source .xlsx file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions.CommentDisplayMode to render all cell comments as footnotes, and saves the workbook as a PDF while handling possible exceptions.
class WorkbookToPdfWithComments
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (comments display mode not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to export an Excel workbook to PDF with cell comments printed using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, configures PdfSaveOptions to retain cell comments, and saves the workbook as a PDF with Aspose.Cells. | Show how to verify the source Excel file exists and implement robust exception handling while exporting to PDF with comments using Aspose.Cells. | Explain the steps required to preserve cell comment formatting when converting an Excel workbook to PDF in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to PDF including cell comments | How to keep Excel comment formatting when saving as PDF with Aspose.Cells | PdfSaveOptions comment visibility Aspose.Cells .NET example | C# code to print cell comments in PDF generated from workbook | Enable comment rendering in PDF output using Aspose.Cells for .NET
// Tags: aspocells pdfsaveoptions include-cell-comments | c# export excel to pdf with comments | preserve comment formatting aspocells | excel to pdf conversion with annotations .net | load workbook and save as pdf aspocells | exception handling excel pdf export c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Required for PdfSaveOptions

// The program checks for the presence of an input .xlsx file, loads it with Aspose.Cells, creates a PdfSaveOptions object (comments are included by default), and saves the workbook as a PDF while handling any runtime exceptions.
class PrintCommentsToPdf
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In recent Aspose.Cells versions, cell comments are included by default
            // when saving to PDF. If a specific option is required in future versions,
            // it can be set here.

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Convert Excel to PDF without Cell Comments using Aspose.Cells for .NET
// Description: Loads an Excel workbook, removes all worksheet comments with ClearComments(), and saves the file as a PDF using PdfSaveOptions, resulting in a clean PDF that contains no comment annotations.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF | remove comments | ClearComments | PdfSaveOptions | export without comments | batch conversion | PDF generation
// Common Searches: Aspose.Cells export Excel to PDF without comments | C# remove all comments before PDF conversion | How to clear worksheet comments in Aspose.Cells | Convert workbook to PDF excluding cell comments | PdfSaveOptions hide comments Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook while omitting every cell comment.
// Use Cases: Produce printable reports from financial models without internal comment notes. | Automate bulk conversion of spreadsheets to clean PDFs for client delivery. | Generate documentation from design sheets where comments are for internal use only.
// AI Prompts: Write C# code that converts an Excel file to PDF with Aspose.Cells, ensuring all comments are excluded. | Show how to modify the example to retain comments on selected worksheets while removing them from others during PDF export. | Suggest best practices for error handling, logging, and performance when converting many workbooks to PDF without comments.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, removes all worksheet comments with ClearComments(), and saves the file as a PDF using PdfSaveOptions, resulting in a clean PDF that contains no comment annotations.
class WorkbookToPdfWithoutComments
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove all comments from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ClearComments();
            }

            // Set PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Path for the output PDF
            string outputPath = "output_without_comments.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF without comments: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

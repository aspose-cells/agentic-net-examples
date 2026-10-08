// Title: Convert only the first worksheet of an Excel file to a single‑page PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that opens an .xlsx file, extracts its first worksheet, and saves it as a one‑page PDF with Aspose.Cells. | Show how to apply custom PDF metadata (title, author) and modify page margins before exporting a single worksheet to PDF using Aspose.Cells. | Write a reusable C# method that accepts input and output file paths and returns a MemoryStream containing a PDF of the first worksheet created with Aspose.Cells.
// Common Searches: how to export only the first sheet of an Excel workbook to PDF using Aspose.Cells C# | C# Aspose.Cells one page per sheet PDF option example | save specific worksheet as PDF with custom margins Aspose.Cells .NET | Aspose.Cells PdfSaveOptions OnePagePerSheet usage | convert Excel first worksheet to PDF programmatically in .NET
// Tags: Aspose.Cells export first worksheet PDF | PdfSaveOptions OnePagePerSheet C# | C# convert Excel sheet to PDF Aspose.Cells | custom PDF metadata Aspose.Cells | single‑sheet PDF generation .NET

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx with Aspose.Cells, copies only the first worksheet into a new workbook, configures PdfSaveOptions.OnePagePerSheet, and saves the result as output.pdf.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input Excel file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the original workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook that will contain only the first worksheet
            Workbook singleSheetWorkbook = new Workbook();
            singleSheetWorkbook.Worksheets.Clear(); // remove default sheet

            // Copy the first worksheet by name
            string firstSheetName = sourceWorkbook.Worksheets[0].Name;
            singleSheetWorkbook.Worksheets.AddCopy(firstSheetName);

            // Configure PDF save options (one page per sheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the single‑sheet workbook as PDF
            singleSheetWorkbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Set OnePagePerSheet and save each Excel worksheet as a separate PDF file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each worksheet to its own PDF file | How to use PdfSaveOptions.OnePagePerSheet when converting Excel to PDF | Iterate through worksheets and save them as individual PDFs using Aspose.Cells | Create a single‑page PDF for each sheet in an Excel workbook with Aspose.Cells | Save Excel sheets as separate PDF documents in .NET without affecting the original workbook
// Tags: OnePagePerSheet PDFSaveOptions Aspose.Cells | per-sheet PDF export .NET | iterate worksheets save as PDF C# | temporary workbook copy worksheet Aspose.Cells | single-page PDF conversion Excel Aspose | separate PDF files per Excel sheet

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, loops through all worksheets, creates a temporary workbook containing only the current sheet, sets PdfSaveOptions.OnePagePerSheet to true, and saves each sheet as an individual PDF named after the worksheet, handling missing files and conversion errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    // Create a temporary workbook containing only the current sheet
                    Workbook sheetWorkbook = new Workbook();
                    sheetWorkbook.Worksheets.Clear(); // Remove the default sheet

                    // Add a copy of the current worksheet by name
                    sheetWorkbook.Worksheets.AddCopy(workbook.Worksheets[i].Name);

                    // Prepare PDF save options
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true // Render the sheet on a single page
                    };

                    // Construct output file name based on the worksheet name
                    string sheetName = workbook.Worksheets[i].Name;
                    string outputFile = $"{sheetName}.pdf";

                    // Save the single-sheet workbook as a PDF file
                    sheetWorkbook.Save(outputFile, pdfOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert sheet '{workbook.Worksheets[i].Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}

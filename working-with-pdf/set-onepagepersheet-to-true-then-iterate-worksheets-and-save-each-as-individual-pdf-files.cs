// Title: How to set OnePagePerSheet and export each Excel worksheet as a separate PDF file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, iterates through all worksheets, creates a fresh Workbook for each sheet, configures PdfSaveOptions to generate a single page per worksheet, and saves the result as a PDF with a safe file name. | Generate a method that copies a given worksheet into a new temporary workbook, applies PDF export settings for one‑page output, and returns the path of the created PDF file. | Provide per‑sheet error handling that logs any exception during PDF conversion and continues processing the remaining worksheets.
// Common Searches: aspnet convert each Excel worksheet to separate PDF file using Aspose.Cells OnePagePerSheet | c# loop through worksheets and export to single-page PDFs with Aspose.Cells | how to sanitize worksheet names for PDF filenames when exporting Excel sheets with Aspose.Cells | batch export Excel sheets to PDFs one page per sheet Aspose.Cells .NET example
// Tags: Aspose.Cells single-page PDF export | PDF export per worksheet Aspose.Cells | worksheet name sanitization for file output | batch PDF generation from Excel .NET | PdfSaveOptions configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel workbook, iterates each worksheet, copies it into a temporary Workbook, sets PdfSaveOptions to fit the sheet on a single page, sanitizes the sheet name for a safe filename, and saves each sheet as an individual PDF while handling errors per sheet.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and export it as an individual PDF
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                try
                {
                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();
                    // Remove the default sheet that comes with a new workbook
                    tempWb.Worksheets.Clear();

                    // Add a new blank worksheet to the temporary workbook
                    Worksheet tempSheet = tempWb.Worksheets.Add(sheet.Name);
                    // Copy the current sheet's content into the temporary worksheet
                    sheet.Copy(tempSheet);

                    // Configure PDF save options to fit the sheet on a single page
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true
                    };

                    // Build the output PDF file name
                    string safeSheetName = string.Concat(sheet.Name.Split(Path.GetInvalidFileNameChars()));
                    string outputFile = $"Sheet_{i + 1}_{safeSheetName}.pdf";

                    // Save the temporary workbook as a PDF
                    tempWb.Save(outputFile, pdfOptions);
                    Console.WriteLine($"Saved: {outputFile}");
                }
                catch (Exception ex)
                {
                    // Handle errors for the current worksheet without stopping the whole process
                    Console.WriteLine($"Failed to process sheet \"{sheet.Name}\": {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

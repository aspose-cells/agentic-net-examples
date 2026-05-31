using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Hide rows 5 through 10 (0‑based index 4, count 6)
            workbook.Worksheets[0].Cells.HideRows(4, 6);

            // Prepare PDF save options (default settings are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the modified workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime errors and display a concise message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
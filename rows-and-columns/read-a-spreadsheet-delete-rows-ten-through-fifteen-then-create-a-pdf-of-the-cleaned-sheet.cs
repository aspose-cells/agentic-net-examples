using System;
using System.IO;
using Aspose.Cells;

class DeleteRowsAndExportPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = @"C:\Path\To\Your\InputFile.xlsx";

        // Path where the resulting PDF will be saved
        string outputPdf = @"C:\Path\To\Your\CleanedOutput.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete rows 10 through 15 (zero‑based index: start at 9, delete 6 rows)
            worksheet.Cells.DeleteRows(9, 6);

            // Save the modified workbook as PDF
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to: {outputPdf}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
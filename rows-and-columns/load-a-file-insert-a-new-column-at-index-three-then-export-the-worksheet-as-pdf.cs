using System;
using System.IO;
using Aspose.Cells;

class InsertColumnAndExportPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a new column at index 3 (fourth column, zero‑based) and update references
            sheet.Cells.InsertColumn(3, true);

            // Optional PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Path for the output PDF file
            string outputPath = "output.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine("Column inserted and worksheet exported to PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
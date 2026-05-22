using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // PdfSaveOptions

class ExportFilteredPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "filtered_output.pdf";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide rows 50‑55 (zero‑based index: start at 49, hide 6 rows)
            workbook.Worksheets[0].Cells.HideRows(49, 6);

            // Set PDF save options (default behavior respects hidden rows)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the visible portion as PDF
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
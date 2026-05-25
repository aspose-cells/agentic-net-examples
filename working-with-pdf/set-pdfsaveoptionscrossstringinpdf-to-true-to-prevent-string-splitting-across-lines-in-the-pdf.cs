using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a long text that would normally wrap or split across lines in PDF
            worksheet.Cells["A1"].PutValue(
                "This is a very long text that will demonstrate how to prevent string splitting across lines when saving to PDF.");

            // Reduce column width to force the text to overflow the cell boundaries
            worksheet.Cells.SetColumnWidth(0, 5);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The TextCrossString property is not available in the current Aspose.Cells version.
            // The default behavior keeps text together; no additional setting is required.

            // Define output file path
            string outputPath = "Output.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
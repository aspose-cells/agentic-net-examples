using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class CsvToPdfConverter
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Import CSV data (comma delimiter, convert numeric values, start at A1)
        worksheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Configure PDF save options to keep all columns on a single page and preserve layout
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,                     // One page per sheet
            AllColumnsInOnePagePerSheet = true,         // Fit all columns on that page
            ExportDocumentStructure = true              // Preserve document structure
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
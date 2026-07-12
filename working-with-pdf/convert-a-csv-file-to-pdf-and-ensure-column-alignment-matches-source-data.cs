using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // PdfSaveOptions

// Author: Aspose.Cells .NET example – CSV to PDF with column alignment
class CsvToPdfConverter
{
    static void Main()
    {
        // Paths to the source CSV and the target PDF
        string csvPath = "input.csv";
        string pdfPath = "output.pdf";

        // Load the CSV file. LoadOptions with LoadFormat.Csv ensures proper parsing.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Auto‑fit columns so the PDF reflects the exact column widths of the CSV data.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AutoFitColumns();

        // Configure PDF save options.
        // AllColumnsInOnePagePerSheet keeps every column on a single page,
        // preserving the visual alignment of the source data.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a PDF using the configured options.
        workbook.Save(pdfPath, pdfOptions);
    }
}
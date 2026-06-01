using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

class HideColumnsAndExportPdf
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns K (index 10) through M (index 12) – total of 3 columns
        int startColumn = 10; // Column K (zero‑based)
        int columnCount = 3;  // K, L, M
        cells.HideColumns(startColumn, columnCount);

        // Set PDF save options (default options will include hidden columns in the output)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF; hidden columns will appear in the PDF
        string outputPath = "output.pdf";
        workbook.Save(outputPath, pdfOptions);
    }
}
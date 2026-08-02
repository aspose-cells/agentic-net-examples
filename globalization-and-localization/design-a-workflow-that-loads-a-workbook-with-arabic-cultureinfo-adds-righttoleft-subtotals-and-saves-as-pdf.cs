using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Saving;

class ArabicRtlSubtotalPdf
{
    static void Main()
    {
        // Set Arabic culture for the current thread (affects number/date formatting)
        CultureInfo arabicCulture = new CultureInfo("ar-SA");
        System.Threading.Thread.CurrentThread.CurrentCulture = arabicCulture;
        System.Threading.Thread.CurrentThread.CurrentUICulture = arabicCulture;

        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for Arabic layout
        sheet.DisplayRightToLeft = true;

        // Define the data range for subtotals (example: A1:C6)
        CellArea dataRange = CellArea.CreateCellArea(0, 0, 5, 2); // rows 0‑5, columns 0‑2

        // Add subtotals:
        // - Group by the first column (index 0)
        // - Use SUM function on the third column (index 2)
        // - Replace existing subtotals, no page breaks, summary placed below data
        sheet.Cells.Subtotal(
            dataRange,
            groupBy: 0,
            function: ConsolidationFunction.Sum,
            totalList: new int[] { 2 },
            replace: true,
            pageBreaks: false,
            summaryBelowData: true);

        // Configure PDF save options suitable for Arabic text
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "Arial",               // Arabic‑compatible font
            CheckWorkbookDefaultFont = true      // Ensure default font is used when needed
        };

        // Save the workbook as PDF
        string outputPath = "output.pdf";
        workbook.Save(outputPath, pdfOptions);
    }
}
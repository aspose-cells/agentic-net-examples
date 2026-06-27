using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SubtotalItalianPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Load the workbook with Italian culture settings
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("it-IT"); // Italian CultureInfo
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the range that contains the data (e.g., A1:C6)
        // Adjust the range as needed for your actual data
        CellArea dataArea = CellArea.CreateCellArea(0, 0, 5, 2); // rows 0-5, columns 0-2

        // Add subtotals:
        // - Group by the first column (index 0)
        // - Use SUM function
        // - Apply subtotal to the third column (index 2)
        // - Replace existing subtotals, add page breaks, place summary below data
        cells.Subtotal(
            dataArea,
            0,
            ConsolidationFunction.Sum,
            new int[] { 2 },
            replace: true,
            pageBreaks: true,
            summaryBelowData: true);

        // Save the workbook as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save(outputPath, pdfOptions);
    }
}
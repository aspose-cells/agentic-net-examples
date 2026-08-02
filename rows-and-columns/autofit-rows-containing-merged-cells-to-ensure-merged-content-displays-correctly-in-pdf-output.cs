using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AutoFitMergedCellsPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range that spans A1:B2 and merge it.
            AsposeRange mergedRange = worksheet.Cells.CreateRange(0, 0, 2, 2);
            mergedRange.Merge();

            // Put a long text into the merged cell and enable text wrapping.
            worksheet.Cells[0, 0].PutValue(
                "This is a long piece of text that should wrap inside the merged cells and demonstrate auto‑fitting of row height when exporting to PDF.");
            Style style = worksheet.Cells[0, 0].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells[0, 0].SetStyle(style);

            // Configure AutoFitterOptions to auto‑fit merged cells.
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine
            };

            // Apply the auto‑fit operation with the specified options.
            worksheet.AutoFitRows(options);

            // Save the workbook as PDF.
            workbook.Save("AutoFitMergedCells.pdf", SaveFormat.Pdf);
        }
    }
}
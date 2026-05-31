using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitMergedRowsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell and enable text wrapping
            worksheet.Cells["A1"].PutValue("This is a long sample text that will be placed inside merged cells to demonstrate auto‑fitting of row heights when the cells are merged.");
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Merge a range of cells (A1:B3) – rows 0‑2 and columns 0‑1
            worksheet.Cells.Merge(0, 0, 3, 2); // Merge A1:B3

            // Configure AutoFitterOptions to consider merged cells.
            // AutoFitMergedCellsType.EachLine expands the height of every row in the merged area.
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                // Optional: improve wrapped‑text handling
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // Auto‑fit all rows in the worksheet using the specified options
            worksheet.AutoFitRows(options);

            // Save the workbook to a file
            workbook.Save("AutoFitMergedRowsDemo.xlsx");
        }
    }
}
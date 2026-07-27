using System;
using Aspose.Cells;

namespace AutoFitMergedCellsDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into a cell and merge a range of cells (A1:B3)
            sheet.Cells["A1"].PutValue("This is a sample text for merged cells auto‑fit demonstration. It should wrap and cause the row height to adjust based on the selected AutoFitMergedCellsType.");
            sheet.Cells.Merge(0, 0, 3, 2); // Merge A1:B3

            // Enable text wrapping for the merged cell
            Style style = sheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to expand the height of each row in the merged range
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // Auto‑fit all rows in the worksheet using the specified options
            sheet.AutoFitRows(options);

            // (Optional) Auto‑fit columns as well, using the same merged‑cell handling
            sheet.AutoFitColumns(options);

            // Save the workbook to a file
            workbook.Save("AutoFitMergedCellsDemo.xlsx");
        }
    }
}
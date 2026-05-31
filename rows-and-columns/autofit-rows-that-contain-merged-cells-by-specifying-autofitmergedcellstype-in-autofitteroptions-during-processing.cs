using System;
using Aspose.Cells;

namespace AutoFitMergedCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into the top‑left cell
            worksheet.Cells["A1"].PutValue("This is a sample text for merged cells auto‑fit demonstration. " +
                                          "It is intentionally long to require row height adjustment when the cells are merged.");

            // Merge a range of cells (A1:B3)
            worksheet.Cells.Merge(0, 0, 3, 2); // rows 0‑2, columns 0‑1

            // Enable text wrapping so the content can span multiple lines
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Create AutoFitterOptions and set the merged‑cells handling type
            AutoFitterOptions options = new AutoFitterOptions
            {
                // Expand the height of each row that participates in the merged area
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                // Optional: treat wrapped text as paragraphs for better height calculation
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // Auto‑fit all rows in the worksheet using the specified options
            worksheet.AutoFitRows(options);

            // Save the workbook to a file
            workbook.Save("AutoFitMergedCellsDemo.xlsx");
        }
    }
}
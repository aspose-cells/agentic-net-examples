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
            sheet.Cells["A1"].PutValue("This is a sample text that is long enough to require row height adjustment when the cells are merged.");
            sheet.Cells.Merge(0, 0, 3, 2); // Merge rows 0-2 and columns 0-1 (A1:B3)

            // Enable text wrapping for the merged cell
            Style style = sheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to expand the height of each row in merged cells
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                // Optional: also fit wrapped text by paragraph
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // AutoFit all rows in the worksheet using the configured options
            sheet.AutoFitRows(options);

            // Save the workbook
            workbook.Save("AutoFitMergedCellsResult.xlsx");
        }
    }
}
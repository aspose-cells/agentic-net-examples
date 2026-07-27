using System;
using Aspose.Cells;

namespace MergedCellsPdfDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge a range of cells (rows 0-2, columns 0-1) => 3 rows x 2 columns
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 0, 3, 2);

            // Set a value in the merged cell (upper‑left cell of the range)
            sheet.Cells[0, 0].Value = "Merged Cells in PDF";

            // Apply basic styling to the merged cell
            Style style = sheet.Cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Font.IsBold = true;
            sheet.Cells[0, 0].SetStyle(style);

            // Save the workbook as PDF – merged cells retain their spanning
            workbook.Save("MergedCellsOutput.pdf", SaveFormat.Pdf);
        }
    }
}
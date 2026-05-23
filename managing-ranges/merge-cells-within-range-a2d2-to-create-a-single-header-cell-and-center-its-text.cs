using System;
using Aspose.Cells;

namespace AsposeCellsMergeHeader
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells A2:D2 (row index 1, column index 0, 1 row, 4 columns)
            cells.Merge(1, 0, 1, 4);

            // Set the header text in the merged cell (upper‑left cell of the range)
            cells[1, 0].Value = "Header";

            // Create a style to center the text horizontally and vertically
            Style style = cells[1, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            cells[1, 0].SetStyle(style);

            // Save the workbook
            workbook.Save("MergedHeader.xlsx");
        }
    }
}
using System;
using Aspose.Cells;

namespace MergeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells B2 (row 1, column 1) through D2 (row 1, column 3)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(1, 1, 1, 3);

            // Set a sample value in the merged cell
            cells[1, 1].PutValue("Centered Text");

            // Apply horizontal center alignment to the merged cell
            Style style = cells[1, 1].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            cells[1, 1].SetStyle(style);

            // Save the workbook
            workbook.Save("MergedCells_B2_D2.xlsx");
        }
    }
}
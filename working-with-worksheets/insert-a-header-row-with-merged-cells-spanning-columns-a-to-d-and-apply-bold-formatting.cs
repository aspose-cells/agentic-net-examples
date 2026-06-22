using System;
using Aspose.Cells;

namespace AsposeCellsHeaderExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set the header text in the first cell (A1)
            cells["A1"].PutValue("Header Title");

            // Merge cells from A1 to D1 (columns 0 to 3, single row)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(0, 0, 1, 4);

            // Apply bold formatting to the merged header cell
            Style headerStyle = cells["A1"].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            cells["A1"].SetStyle(headerStyle);

            // Save the workbook to a file
            workbook.Save("HeaderWithMergedCells.xlsx");
        }
    }
}
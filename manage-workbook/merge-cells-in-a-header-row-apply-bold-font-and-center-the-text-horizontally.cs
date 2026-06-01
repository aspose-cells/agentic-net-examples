using System;
using Aspose.Cells;

namespace AsposeCellsHeaderMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells in the first row (A1:D1) to create a header area
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(0, 0, 1, 4);

            // Set the header text in the merged cell (upper‑left cell of the range)
            cells[0, 0].PutValue("Report Header");

            // Retrieve the style of the merged cell
            Style headerStyle = cells[0, 0].GetStyle();

            // Apply bold font
            headerStyle.Font.IsBold = true;

            // Center the text horizontally
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Apply the modified style back to the merged cell
            cells[0, 0].SetStyle(headerStyle);

            // Save the workbook to a file
            workbook.Save("HeaderMerged.xlsx");
        }
    }
}
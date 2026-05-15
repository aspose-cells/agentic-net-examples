using System;
using Aspose.Cells;

namespace AsposeCellsMergeHeader
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells in the first row (A1:D1)
            cells.Merge(0, 0, 1, 4); // firstRow, firstColumn, totalRows, totalColumns

            // Set the header text in the merged cell
            cells[0, 0].PutValue("Report Header");

            // Retrieve the style of the merged cell
            Style style = cells[0, 0].GetStyle();

            // Apply bold font
            style.Font.IsBold = true;

            // Center the text horizontally
            style.HorizontalAlignment = TextAlignmentType.Center;

            // Apply the modified style back to the merged cell
            cells[0, 0].SetStyle(style);

            // Save the workbook
            workbook.Save("HeaderMerged.xlsx");
        }
    }
}
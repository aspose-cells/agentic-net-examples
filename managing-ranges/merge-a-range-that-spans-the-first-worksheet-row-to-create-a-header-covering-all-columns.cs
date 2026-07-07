using System;
using Aspose.Cells;

class MergeHeaderDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define how many columns the header should span (e.g., 10 columns)
        int totalColumns = 10;

        // Merge the first row (row index 0) across the defined columns
        // Parameters: firstRow, firstColumn, totalRows (1), totalColumns
        cells.Merge(0, 0, 1, totalColumns);

        // Set the header text in the merged cell
        cells[0, 0].PutValue("Report Header");

        // Apply a simple style: centered and bold
        Style style = cells[0, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        cells[0, 0].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("HeaderMerged.xlsx");
    }
}
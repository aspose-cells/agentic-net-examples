using System;
using Aspose.Cells;

class MergeAndCenterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells D4 through F4 (zero‑based indices: row 3, column 3, 1 row, 3 columns)
        worksheet.Cells.Merge(3, 3, 1, 3);

        // Optional: put a value in the merged cell
        worksheet.Cells[3, 3].PutValue("Merged and Centered");

        // Retrieve the current style of the merged cell
        Style style = worksheet.Cells[3, 3].GetStyle();

        // Center the content horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the modified style back to the merged cell
        worksheet.Cells[3, 3].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedCentered.xlsx");
    }
}
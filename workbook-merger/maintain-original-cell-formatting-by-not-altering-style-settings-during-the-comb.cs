using System;
using Aspose.Cells;

class PreserveMergeStyleDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the top‑left cell of the range to be merged
        int startRow = 1;   // zero‑based index (row 2 in Excel)
        int startCol = 1;   // zero‑based index (column B)

        // Put a value and apply a custom style to this cell
        cells[startRow, startCol].PutValue("Header");
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = System.Drawing.Color.White;
        headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        headerStyle.VerticalAlignment = TextAlignmentType.Center;
        cells[startRow, startCol].SetStyle(headerStyle);

        // Capture the original style before merging
        Style originalStyle = cells[startRow, startCol].GetStyle();

        // Merge a 3‑row by 2‑column block (rows 2‑4, columns B‑C)
        cells.Merge(startRow, startCol, 3, 2);

        // Re‑apply the captured style to the merged cell.
        // The explicitFlag set to true ensures only explicitly set properties are overwritten,
        // preserving any other formatting that might have been altered by the merge.
        cells[startRow, startCol].SetStyle(originalStyle, true);

        // Save the workbook
        workbook.Save("PreserveMergeStyle.xlsx");
    }
}
using System;
using Aspose.Cells;
using System.Drawing;

class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the header range (e.g., A1 to D1)
        int firstRow = 0;          // zero‑based index for the first row (header row)
        int firstColumn = 0;       // zero‑based index for the first column (A)
        int totalRows = 1;         // only one row for the header
        int totalColumns = 4;      // span four columns (A to D)

        // Merge the header cells into a single cell
        cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

        // Create a style with navy blue fill color
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.ForegroundColor = Color.Navy;

        // Apply the style to the merged cell (top‑left cell of the range)
        cells[firstRow, firstColumn].SetStyle(headerStyle);

        // Optional: set a header text
        cells[firstRow, firstColumn].PutValue("Header");

        // Save the workbook
        workbook.Save("MergedHeader.xlsx");
    }
}
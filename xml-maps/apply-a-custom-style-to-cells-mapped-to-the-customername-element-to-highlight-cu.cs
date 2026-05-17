using Aspose.Cells;
using System.Drawing;

class HighlightCustomerNames
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a custom style to highlight customer names
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;          // background color
        highlightStyle.Pattern = BackgroundType.Solid;          // solid fill
        highlightStyle.Font.IsBold = true;                      // bold font

        // Assume the /Customer/Name values are placed in column B (index 1) starting from row 2
        int nameColumnIndex = 1; // column B (zero‑based)
        int startRow = 1;        // row 2 (zero‑based)
        int endRow = worksheet.Cells.MaxDataRow; // last row with data

        // Apply the custom style to each non‑empty cell in the name column
        for (int row = startRow; row <= endRow; row++)
        {
            Cell cell = worksheet.Cells[row, nameColumnIndex];
            if (cell.Value != null)
            {
                cell.SetStyle(highlightStyle);
            }
        }

        // Save the workbook with the applied style
        workbook.Save("output.xlsx");
    }
}
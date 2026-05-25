using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cells collection
        Cells cells = worksheet.Cells;

        // Specify the zero‑based row and column indices
        int rowIndex = 2;      // third row (A3)
        int columnIndex = 3;   // fourth column (D)

        // Access the cell by numeric index (Cells indexer rule)
        Cell cell = cells[rowIndex, columnIndex];

        // Optional: put a value into the cell
        cell.PutValue("Demo");

        // Retrieve the cell's current style
        Style style = cell.GetStyle();

        // Set a solid background color (using ForegroundColor with solid pattern)
        style.ForegroundColor = Color.Yellow;
        style.Pattern = BackgroundType.Solid;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook (save rule)
        workbook.Save("ModifiedCell.xlsx");
    }
}
using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection
            Cells cells = worksheet.Cells;

            // Access a cell by numeric row and column indexes (zero‑based)
            // Example: row 2, column 3 corresponds to cell D3
            Cell cell = cells[2, 3];

            // Optional: put a value into the cell
            cell.PutValue("Demo");

            // Create a style and set its background color
            Style style = workbook.CreateStyle();
            style.BackgroundColor = Color.LightBlue;   // Desired background color
            style.Pattern = BackgroundType.Solid;      // Ensure the color is applied

            // Apply the style to the selected cell
            cell.SetStyle(style);

            // Save the workbook to a file
            workbook.Save("CellBackgroundDemo.xlsx");
        }
    }
}
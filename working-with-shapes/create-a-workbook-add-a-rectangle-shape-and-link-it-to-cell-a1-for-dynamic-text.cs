// Title: Create an Excel workbook and add a rectangle shape whose text is linked to cell A1 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a new workbook, insert a rectangle shape at a specific position, and set its Text property to "=A1" so the shape displays the cell value dynamically. | Write C# code that creates a worksheet, puts "Dynamic Text" into A1, adds a rectangle shape, and binds the shape’s displayed text to the cell using Aspose.Cells. | Save the workbook as an .xlsx file after linking the rectangle shape’s text to cell A1 with a formula.
// Common Searches: Aspose.Cells C# how to bind a shape's text to a worksheet cell value | link rectangle shape text to cell A1 in Excel using Aspose.Cells .NET | dynamic shape text from cell formula Aspose.Cells example | add rectangle shape programmatically and display cell content in C#
// Tags: add rectangle shape Aspose.Cells C# | bind shape text to cell formula Aspose.Cells | set shape Text property =A1 Excel | create workbook with linked shape Aspose.Cells | dynamic shape text from worksheet cell C#

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes "Dynamic Text" to cell A1, adds a rectangle shape at row 2/column 2 with a 100 × 50 pixel size, sets the shape’s Text property to the formula "=A1" so it reflects the cell’s value dynamically, and saves the file as RectangleLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put some initial text in cell A1
        sheet.Cells["A1"].PutValue("Dynamic Text");

        // Add a rectangle shape.
        // Parameters: shape type, upper left row, upper left column, top, left, bottom, right (all in pixels)
        // Position it starting at row 2 (index 1), column 2 (index 1) with a size of 100x50 pixels.
        int upperLeftRow = 1;      // Row index (zero‑based)
        int upperLeftColumn = 1;   // Column index (zero‑based)
        int top = 5;               // Pixels from the top of the cell
        int left = 5;              // Pixels from the left of the cell
        int bottom = 55;           // Pixels from the top (height = bottom - top)
        int right = 105;           // Pixels from the left (width = right - left)

        Shape rectangle = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle,
            upperLeftRow,
            upperLeftColumn,
            top,
            left,
            bottom,
            right);

        // Link the rectangle's displayed text to cell A1.
        // Setting the Text property to a formula (e.g., "=A1") makes the shape show the cell's value dynamically.
        rectangle.Text = "=A1";

        // Save the workbook to a file
        workbook.Save("RectangleLinkedToCell.xlsx");
    }
}

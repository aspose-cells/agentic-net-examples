// Title: Link a Rectangle Shape to a Cell Using LEFT & RIGHT Functions – Aspose.Cells for .NET
// Description: Creates a workbook, writes "AsposeCellsDemo" to A1, sets B1 to =LEFT(A1,5)&RIGHT(A1,4), adds a rectangle shape, links the shape to B1 with the LinkedCell property, and saves the file so the shape shows the concatenated substring result.
// Keywords: Aspose.Cells | C# | shape linked cell | LEFT function | RIGHT function | rectangle shape | LinkedCell property | Excel automation | substring formula | display shape text
// Common Searches: Aspose.Cells link shape to cell .NET | How to use LEFT and RIGHT in a cell formula with Aspose.Cells | Set LinkedCell for a rectangle shape in C# | Show formula result in a linked shape using Aspose.Cells | Create dynamic text in Excel shapes via Aspose.Cells
// Developer Intent: Attach a rectangle shape to a cell that contains a LEFT/RIGHT substring formula so the shape automatically displays the computed text.
// Use Cases: Generate a report where a shape reflects a custom string derived from another cell. | Build dashboards with shapes that update instantly when the source cell value changes. | Automate Excel templates that require visual labels derived from cell formulas.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape linked to a cell that uses LEFT and RIGHT functions. | Show how to modify the linked cell formula to extract different parts of a string and have the shape update automatically. | Explain how to read the text displayed by a shape linked to a formula cell after the workbook is saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes "AsposeCellsDemo" to A1, sets B1 to =LEFT(A1,5)&RIGHT(A1,4), adds a rectangle shape, links the shape to B1 with the LinkedCell property, and saves the file so the shape shows the concatenated substring result.
class ShapeLinkedCellExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put some sample text in cell A1
        sheet.Cells["A1"].PutValue("AsposeCellsDemo");

        // In cell B1 place a formula that extracts substrings using LEFT and RIGHT
        // Example: first 5 characters and last 4 characters concatenated
        sheet.Cells["B1"].Formula = "=LEFT(A1,5) & RIGHT(A1,4)";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset (pixels), 
        // lower right row, lower right column, lower right offset (pixels)
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 2, 5, 0);

        // Link the shape to the cell containing the formula (B1)
        // The shape will display the result of the formula
        shape.LinkedCell = "$B$1";

        // Optionally, set the shape's text to show the linked value (not required,
        // as the linked cell value is displayed automatically in Excel)
        // shape.Text = sheet.Cells["B1"].StringValue; // Uncomment if needed

        // Save the workbook
        workbook.Save("ShapeLinkedCellExample.xlsx");
    }
}

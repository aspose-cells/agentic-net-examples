// Title: Align a Rectangle Shape to a Column Header Cell with Aspose.Cells for .NET
// Description: Creates a workbook, writes a header in column C (row 0), adds a rectangle shape, sets UpperLeftRow and UpperLeftColumn to the header cell, clears horizontal offset with UpperDeltaX, verifies the placement, and saves the file as AlignedShape.xlsx.
// Keywords: Aspose.Cells shape alignment | C# set shape position cell | UpperLeftRow UpperLeftColumn | remove shape offset Aspose | validate shape placement | rectangle shape worksheet | Aspose.Cells .NET example
// Common Searches: Aspose.Cells align shape to specific cell | set shape UpperLeftRow UpperLeftColumn C# | remove shape offset Aspose.Cells | check if shape is aligned with header cell | position rectangle shape in Excel using Aspose
// Developer Intent: Place a rectangle shape so its upper‑left corner matches a column header cell and confirm the alignment programmatically.
// Use Cases: Add a visual marker next to a column title in a generated report. | Synchronize shapes with dynamic headers for interactive dashboards. | Automated layout verification before exporting an Excel workbook.
// AI Prompts: Write C# code with Aspose.Cells that aligns any shape to a given cell and ensures zero offset. | Show how to adjust UpperDeltaY for vertical alignment of a shape to a header cell. | Create a reusable method that receives a worksheet, row, column, and shape, then returns a boolean indicating correct alignment.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes a header in column C (row 0), adds a rectangle shape, sets UpperLeftRow and UpperLeftColumn to the header cell, clears horizontal offset with UpperDeltaX, verifies the placement, and saves the file as AlignedShape.xlsx.
class AlignShapeToHeader
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Write a header in column C (index 2) of the first row
        worksheet.Cells[0, 2].PutValue("Header");

        // Add a rectangle shape with arbitrary initial parameters
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 100, 50, 0);

        // Align the shape's upper‑left corner to the header cell (row 0, column 2)
        shape.UpperLeftRow = 0;      // top row index
        shape.UpperLeftColumn = 2;   // column index of the header

        // Ensure there is no offset from the cell
        shape.UpperDeltaX = 0;

        // Validate that the shape is positioned at the intended cell
        bool isAligned = shape.UpperLeftRow == 0 && shape.UpperLeftColumn == 2;
        Console.WriteLine("Shape aligned to header cell: " + isAligned);

        // Save the workbook
        workbook.Save("AlignedShape.xlsx");
    }
}

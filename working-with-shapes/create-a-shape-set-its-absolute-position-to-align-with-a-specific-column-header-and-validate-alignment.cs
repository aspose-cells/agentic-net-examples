// Title: Align a Rectangle Shape to a Column Header with Aspose.Cells for .NET
// Description: This C# example creates a workbook, writes a header in cell B1, adds a rectangle shape, positions the shape by setting UpperLeftColumn and UpperLeftRow to the header's coordinates, clears any pixel offsets, verifies the placement, and saves the file as AlignedShape.xlsx.
// Keywords: Aspose.Cells shape positioning | C# rectangle shape alignment | UpperLeftColumn property | UpperLeftRow property | reset shape pixel offset | Excel shape alignment example | Aspose.Cells .NET tutorial | shape placement validation
// Common Searches: how to align a shape with a column header using Aspose.Cells | set shape UpperLeftColumn and UpperLeftRow in C# | reset shape pixel offsets Aspose.Cells | verify shape location after moving in Aspose.Cells | Aspose.Cells align rectangle to cell B1
// Developer Intent: Place a rectangle so its top‑left corner matches the exact cell of a column header and programmatically confirm that the coordinates are correct.
// Use Cases: Designing custom report headers where graphics must line up with column titles | Building interactive Excel dashboards that anchor shapes to specific data columns | Automating quality checks for generated spreadsheets to ensure visual elements are correctly positioned
// AI Prompts: Write C# code with Aspose.Cells that adds a shape and aligns its UpperLeftColumn/UpperLeftRow to a given header cell, then confirms the alignment. | Show how to move an existing shape to cell C3, clearing UpperDeltaX and UpperDeltaY for precise placement. | Explain the steps to programmatically validate that a shape’s column and row indices correspond to a target cell in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, writes a header in cell B1, adds a rectangle shape, positions the shape by setting UpperLeftColumn and UpperLeftRow to the header's coordinates, clears any pixel offsets, verifies the placement, and saves the file as AlignedShape.xlsx.
class AlignShapeToColumnHeader
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Write a column header in cell B1 (row 0, column 1)
        worksheet.Cells[0, 1].PutValue("Header");

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixel, upperLeftPixel2, height, width
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 100);

        // Align the shape's upper‑left corner with the column header (column B, row 1)
        shape.UpperLeftColumn = 1; // Column index (0‑based)
        shape.UpperLeftRow = 0;    // Row index (0‑based)

        // Optional: reset any pixel offsets
        shape.UpperDeltaX = 0;
        shape.UpperDeltaY = 0; // UpperDeltaY property exists similarly; set to 0 for exact alignment

        // Validate that the shape is aligned with the intended column header
        if (shape.UpperLeftColumn == 1 && shape.UpperLeftRow == 0)
        {
            Console.WriteLine("Shape successfully aligned with the column header.");
        }
        else
        {
            Console.WriteLine("Shape alignment validation failed.");
        }

        // Save the workbook
        workbook.Save("AlignedShape.xlsx");
    }
}

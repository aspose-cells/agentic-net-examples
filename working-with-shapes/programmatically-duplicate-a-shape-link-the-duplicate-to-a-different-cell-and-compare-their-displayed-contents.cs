// Title: C# – Duplicate a Shape, Link to a Different Cell, and Compare Values with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, adds a rectangle shape linked to cell A1, clones the shape to a new location, links the copy to cell B1, retrieves each shape's linked cell address, compares the displayed cell values, checks if the shapes reference the same cell object, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | duplicate shape | AddCopy | SetLinkedCell | linked cell address | compare cell values | rectangle shape | worksheet shapes | example code
// Common Searches: Aspose.Cells duplicate shape C# | How to copy a shape and set a new linked cell in Aspose.Cells | Get linked cell address from a shape Aspose.Cells .NET | Compare values of cells linked to two shapes | AddCopy shape Aspose.Cells example
// Developer Intent: The developer needs to clone an existing shape, bind the clone to a different worksheet cell, and verify that the values displayed by the two linked cells are as expected.
// Use Cases: Design a report template where a header shape is duplicated for sub‑sections, each reflecting a distinct data cell. | Automate financial dashboards that place identical chart placeholders in multiple areas, each linked to separate calculation cells. | Validate workbook layouts by ensuring duplicated shapes reference the correct, unique cells before publishing.
// AI Prompts: Generate C# code using Aspose.Cells to duplicate a shape and link the copy to a specified cell. | Show how to retrieve linked cell addresses from two shapes and compare their values for equality. | Explain the purpose of the parameters in SetLinkedCell when binding a shape to a worksheet cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This Aspose.Cells for .NET example creates a workbook, adds a rectangle shape linked to cell A1, clones the shape to a new location, links the copy to cell B1, retrieves each shape's linked cell address, compares the displayed cell values, checks if the shapes reference the same cell object, and saves the workbook.
class DuplicateShapeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put sample values into two cells that will be linked to the shapes
        sheet.Cells["A1"].PutValue("Original Shape Value");
        sheet.Cells["B1"].PutValue("Duplicate Shape Value");

        // Access the shapes collection of the worksheet
        ShapeCollection shapes = sheet.Shapes;

        // Add a rectangle shape and link it to cell A1
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        RectangleShape originalShape = shapes.AddRectangle(2, 0, 2, 0, 130, 130);
        originalShape.SetLinkedCell("A1", false, false); // link to A1

        // Duplicate the rectangle shape to a new location (row 7, column 7)
        Shape duplicateShape = shapes.AddCopy(originalShape, 7, 0, 7, 0);
        // Link the duplicate shape to a different cell (B1)
        duplicateShape.SetLinkedCell("B1", false, false);

        // Retrieve the linked cell addresses from both shapes
        string originalLinkedCell = originalShape.GetLinkedCell(false, false);
        string duplicateLinkedCell = duplicateShape.GetLinkedCell(false, false);

        // Get the actual Cell objects
        Cell cellOriginal = sheet.Cells[originalLinkedCell];
        Cell cellDuplicate = sheet.Cells[duplicateLinkedCell];

        // Compare the displayed contents (cell values) of the two linked cells
        bool valuesAreEqual = cellOriginal.Value?.ToString() == cellDuplicate.Value?.ToString();

        // Output the comparison result
        Console.WriteLine($"Original shape linked to cell: {originalLinkedCell} with value \"{cellOriginal.Value}\"");
        Console.WriteLine($"Duplicate shape linked to cell: {duplicateLinkedCell} with value \"{cellDuplicate.Value}\"");
        Console.WriteLine($"Are the displayed contents equal? {valuesAreEqual}");

        // Additionally, demonstrate using Cell.Equals to compare the cell references themselves
        bool sameCellReference = cellOriginal.Equals(cellDuplicate);
        Console.WriteLine($"Do both shapes reference the same cell object? {sameCellReference}");

        // Save the workbook to verify the shapes and links (optional)
        workbook.Save("DuplicateShapeExample.xlsx");
    }
}

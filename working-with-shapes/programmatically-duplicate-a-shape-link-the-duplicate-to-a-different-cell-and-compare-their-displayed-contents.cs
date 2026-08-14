// Title: C# – Duplicate a Shape, Link to a New Cell, and Compare Linked Values with Aspose.Cells
// Description: This example creates a workbook, adds a rectangle shape linked to cell A1, copies the shape to another location, links the copy to cell B1, then retrieves each shape's linked cell address, accesses the corresponding Cell objects, and checks both reference equality and displayed values before saving the file.
// Keywords: Aspose.Cells shape copy | SetLinkedCell C# | AddCopy method | compare linked cell values | duplicate shape linked cell | .NET spreadsheet shape manipulation
// Common Searches: Aspose.Cells copy shape and change linked cell | How to compare values of cells linked to shapes in C# | Set different linked cell for shape duplicate Aspose | Retrieve linked cell address from a shape Aspose.Cells
// Developer Intent: Duplicate an existing shape, assign a different linked cell to the copy, and programmatically verify that the two shapes reference distinct cells and display different content.
// Use Cases: Generate a report template where a header shape is cloned for each section, each clone linked to its own title cell. | Create monthly dashboards by copying a chart shape and linking each copy to month‑specific data cells. | Automated testing to ensure shape copies maintain independent links and display the correct cell values.
// AI Prompts: Write C# code with Aspose.Cells that copies a shape, sets a new linked cell, and prints whether the linked cells are the same and have identical values. | Explain the interaction between AddCopy and SetLinkedCell for displaying separate cell contents in duplicated shapes. | Provide a step‑by‑step tutorial on comparing the underlying cell references of two shapes linked to different cells in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDuplicateExample
{
    // This example creates a workbook, adds a rectangle shape linked to cell A1, copies the shape to another location, links the copy to cell B1, then retrieves each shape's linked cell address, accesses the corresponding Cell objects, and checks both reference equality and displayed values before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add a rectangle shape and link it to cell A1
            // -------------------------------------------------
            ShapeCollection shapes = worksheet.Shapes;
            // Add a rectangle shape at row 2, column 2 (zero‑based indexes)
            RectangleShape originalShape = shapes.AddRectangle(2, 0, 2, 0, 130, 130);
            // Link the shape to cell A1
            originalShape.SetLinkedCell("A1", false, false);
            // Put a value into A1 that will be displayed by the shape
            worksheet.Cells["A1"].PutValue("Original");

            // -------------------------------------------------
            // 2. Duplicate the shape and place it at a new location
            // -------------------------------------------------
            // AddCopy creates a copy of the source shape; we move it to row 7, column 7
            Shape duplicatedShape = shapes.AddCopy(originalShape, 7, 0, 7, 0);
            // Link the duplicated shape to a different cell (B1)
            duplicatedShape.SetLinkedCell("B1", false, false);
            // Put a different value into B1
            worksheet.Cells["B1"].PutValue("Duplicate");

            // -------------------------------------------------
            // 3. Retrieve linked cells and compare their displayed contents
            // -------------------------------------------------
            // Get linked cell addresses from both shapes
            string originalLinkedAddress = originalShape.GetLinkedCell(false, false);
            string duplicateLinkedAddress = duplicatedShape.GetLinkedCell(false, false);

            // Obtain the actual Cell objects
            Cell originalCell = worksheet.Cells[originalLinkedAddress];
            Cell duplicateCell = worksheet.Cells[duplicateLinkedAddress];

            // Compare the cell objects (same location?) – should be false
            bool sameCellReference = originalCell.Equals(duplicateCell);
            Console.WriteLine($"Are linked cells the same reference? {sameCellReference}");

            // Compare the displayed values
            bool sameValue = object.Equals(originalCell.Value, duplicateCell.Value);
            Console.WriteLine($"Do linked cells contain the same value? {sameValue}");

            // -------------------------------------------------
            // 4. Save the workbook (optional – demonstrates lifecycle rule)
            // -------------------------------------------------
            workbook.Save("ShapeDuplicateComparison.xlsx");
        }
    }
}

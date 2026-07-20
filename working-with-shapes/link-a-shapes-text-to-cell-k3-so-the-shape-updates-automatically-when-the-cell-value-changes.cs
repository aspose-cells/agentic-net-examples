// Title: C# – Link Shape Text to Cell K3 Using Aspose.Cells (LinkedCell Property)
// Description: Demonstrates how to add a shape to a worksheet with Aspose.Cells for .NET, set its LinkedCell property to the absolute reference "$K$3", and have the shape’s displayed text automatically reflect any changes made to cell K3. The example saves the workbook as LinkedShapeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | LinkedCell | shape text binding | Excel shape to cell | dynamic shape caption | cell K3 | automatic update | worksheet shape | API example
// Common Searches: Aspose.Cells link shape text to cell | C# set shape LinkedCell to K3 | update Excel shape caption from cell value | Aspose.Cells shape text binding example | how to bind a shape to a worksheet cell in .NET
// Developer Intent: Bind a shape’s displayed text to cell K3 so it updates automatically when the cell value changes.
// Use Cases: Create a live dashboard where a shape shows the current total stored in K3 and refreshes with each calculation. | Replace static labels on report graphics with dynamic values by linking each shape to its result cell. | Build interactive templates that automatically propagate formula results into shape captions without manual editing.
// AI Prompts: Write C# code using Aspose.Cells that links a shape’s text to cell K3 and keeps it synchronized with cell changes. | Show how to bind multiple shapes to different cells (e.g., K3, L5, M7) using the LinkedCell property in Aspose.Cells. | Explain the LinkedCell property behavior, supported shape types, and any limitations when linking shape text to worksheet cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a shape to a worksheet with Aspose.Cells for .NET, set its LinkedCell property to the absolute reference "$K$3", and have the shape’s displayed text automatically reflect any changes made to cell K3. The example saves the workbook as LinkedShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (you can choose any shape type)
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 100);

        // Link the shape's text to cell K3 (absolute reference)
        shape.LinkedCell = "$K$3";

        // Optionally set some initial text (will be overridden by the linked cell value)
        shape.Text = "Linked to K3";

        // Save the workbook
        workbook.Save("LinkedShapeDemo.xlsx");
    }
}

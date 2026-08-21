// Title: C# – Add a Rectangle Shape, Link It to a Text‑Formatted Cell, and Display the Formatted Text with Aspose.Cells
// Description: Creates a new workbook, sets cell B2 to Text format (Number format 49), inserts a rectangle shape, links the shape to B2 using an absolute reference, updates the shape to show the cell's string value, retrieves the linked address, and saves the file. Demonstrates how to verify that the shape reflects the formatted text.
// Keywords: Aspose.Cells C# | add rectangle shape | link shape to cell | text number format 49 | display cell value in shape | linked cell address | update shape value | Excel automation Aspose.Cells | shape linked cell absolute reference | Aspose.Cells .NET example
// Common Searches: Aspose.Cells link shape to text formatted cell | C# rectangle shape shows cell value | set cell number format to text and link shape | retrieve linked cell address from shape Aspose.Cells | update shape caption from cell value .NET
// Developer Intent: Find a step‑by‑step C# example that binds a shape to a cell formatted as text and ensures the shape displays the exact formatted string.
// Use Cases: Design a dynamic dashboard where shape captions automatically reflect labels stored in text‑formatted cells. | Generate reports that need shape titles to stay in sync with cell values that use custom text formatting. | Build an Excel template that links multiple shapes to formatted cells for real‑time caption updates.
// AI Prompts: Show me C# code to add a rectangle shape, link it to a text‑formatted cell, and update the shape's displayed value using Aspose.Cells. | How can I retrieve the absolute linked cell address from a shape and read its string value in Aspose.Cells for .NET? | Explain the steps to set a cell's number format to Text (code 49), link a shape to that cell, and verify the shape shows the formatted text.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, sets cell B2 to Text format (Number format 49), inserts a rectangle shape, links the shape to B2 using an absolute reference, updates the shape to show the cell's string value, retrieves the linked address, and saves the file. Demonstrates how to verify that the shape reflects the formatted text.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set cell B2 with text format and a sample value
            Cell cell = sheet.Cells["B2"];
            Style style = cell.GetStyle();
            style.Number = 49; // Text format code
            cell.SetStyle(style);
            cell.PutValue("Hello Aspose!");

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), height (pixels), width (pixels)
            Shape rect = sheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 200);

            // Link the shape to the formatted text cell B2 (use absolute reference)
            rect.SetLinkedCell("B2", true, true);

            // Update the shape's displayed value based on the linked cell
            rect.UpdateSelectedValue();

            // Retrieve the linked cell address (absolute) and its value
            string linkedAddress = rect.GetLinkedCell(true, true); // returns "$B$2"
            // Remove any leading '$' characters to obtain a valid cell name for indexing
            string cleanAddress = linkedAddress.Replace("$", string.Empty);
            string linkedValue = sheet.Cells[cleanAddress].StringValue;

            Console.WriteLine("Shape's linked cell: " + linkedAddress);
            Console.WriteLine("Value in linked cell: " + linkedValue);

            // Save the workbook (optional)
            string outputPath = "ShapeLinkedCellDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

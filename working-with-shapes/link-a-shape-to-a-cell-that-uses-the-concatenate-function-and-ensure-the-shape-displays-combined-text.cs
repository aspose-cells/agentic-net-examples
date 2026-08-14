// Title: Link a Shape to a CONCATENATE Formula Cell in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes text to A1 and B1, sets C1 with a CONCATENATE formula, adds a rectangle shape, links the shape to C1, refreshes the displayed value, fits the shape to the text, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape linking | SetLinkedCell | CONCATENATE formula | rectangle shape | FitToTextSize | UpdateSelectedValue | Excel automation | cell to shape binding
// Common Searches: Aspose.Cells link shape to cell C# | SetLinkedCell with formula Aspose.Cells | Fit shape to text size Aspose.Cells | Refresh shape text after formula change | Rectangle shape bound to CONCATENATE cell
// Developer Intent: The developer needs to bind a rectangle shape to a cell that contains a CONCATENATE formula so the shape automatically displays the combined text.
// Use Cases: Dynamic dashboards where shape captions reflect merged values from several cells. | Printable reports with address blocks inside shapes that update when source cells change. | Interactive worksheets where shape labels stay synchronized with formula‑driven cell content.
// AI Prompts: Provide C# code that links a rectangle shape to a cell with a CONCATENATE formula and updates the shape text using Aspose.Cells. | Show how to bind multiple shapes to different formula cells and automatically fit each shape to its text. | Explain how to ensure a linked shape refreshes when the source cells of a CONCATENATE formula are edited in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes text to A1 and B1, sets C1 with a CONCATENATE formula, adds a rectangle shape, links the shape to C1, refreshes the displayed value, fits the shape to the text, and saves the file using Aspose.Cells for .NET.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells A1 and B1 with sample text
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set cell C1 to concatenate A1 and B1 with a space
            sheet.Cells["C1"].Formula = "=CONCATENATE(A1,\" \",B1)";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
            int upperRow = 2;    // Row index (0‑based)
            int upperCol = 0;    // Column index (0‑based)
            int lowerRow = 5;
            int lowerCol = 3;
            int width = 150;     // width in pixels
            int height = 80;     // height in pixels
            RectangleShape shape = sheet.Shapes.AddRectangle(upperRow, upperCol, lowerRow, lowerCol, width, height);

            // Link the shape to cell C1 so it displays the concatenated result
            shape.SetLinkedCell("C1", true, true);

            // Update the shape's displayed value based on the linked cell
            shape.UpdateSelectedValue();

            // Optionally adjust the shape to fit the text size
            shape.FitToTextSize();

            // Save the workbook to a file
            string outputPath = "ShapeLinkedToConcatenatedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Aspose.Cells for .NET – C# Example: Link a SmartArt (shape) to a worksheet cell
// Description: Demonstrates how to create a workbook, add a rectangle shape as a SmartArt placeholder, set its LinkedCell property to "$C$5" so the shape moves with the cell, and save the file as SmartArtLinkedCell.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel shape linking | LinkedCell property | SmartArt anchor | Shape.LinkToCell | worksheet cell binding | Aspose.Cells example | GitHub code sample | Excel automation
// Common Searches: Aspose.Cells link shape to cell C# | How to set LinkedCell for a shape in Aspose.Cells | SmartArt anchor to worksheet cell using Aspose.Cells | C# example Shape.LinkToCell Aspose.Cells | Excel shape moves with cell Aspose.Cells
// Developer Intent: Attach a shape (used as SmartArt) to a specific cell so the shape follows the cell during row/column changes.
// Use Cases: Create dynamic reports where a SmartArt diagram stays aligned with a data cell. | Design templates with placeholder shapes that are bound to cells for later content updates. | Build dashboards where visual elements move automatically when the underlying data range is resized.
// AI Prompts: Generate C# code with Aspose.Cells that adds a SmartArt shape and links it to cell D10. | Show how to read the current LinkedCell address of a shape and change it to another cell at runtime. | Explain how Shape.LinkedCell affects shape positioning when rows or columns are inserted or deleted.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape as a SmartArt placeholder, set its LinkedCell property to "$C$5" so the shape moves with the cell, and save the file as SmartArtLinkedCell.xlsx using Aspose.Cells for .NET.
class SmartArtLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (as a placeholder for SmartArt) to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height (pixels), width (pixels)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1,   // upper left row (zero‑based)
                1,   // upper left column (zero‑based)
                0,   // upper left row offset (pixels)
                0,   // upper left column offset (pixels)
                300, // height in pixels
                200  // width in pixels
            );

            // Link the shape to cell C5
            shape.LinkedCell = "$C$5";

            // Define output file path
            string outputPath = "SmartArtLinkedCell.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

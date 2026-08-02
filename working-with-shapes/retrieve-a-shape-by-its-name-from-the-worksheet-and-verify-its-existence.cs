// Title: Retrieve a Shape by Name in Aspose.Cells (C#) and Verify Its Presence
// Description: Demonstrates how to add a rectangle to a worksheet, assign a custom name, fetch the shape using the Shapes collection indexer, and confirm its existence with a null‑check before optionally saving the workbook.
// Keywords: Aspose.Cells get shape by name C# | retrieve worksheet shape Aspose.Cells | check shape existence Aspose.Cells | Shapes collection indexer Aspose.Cells | named shape lookup Aspose.Cells
// Common Searches: Aspose.Cells retrieve shape by name | C# find shape in worksheet Aspose.Cells | verify shape exists Aspose.Cells | how to get a named rectangle Aspose.Cells | shape collection indexer example Aspose.Cells
// Developer Intent: Locate a shape using its assigned name and ensure it is available before further processing.
// Use Cases: Modify formatting of a specific chart or image identified by name. | Validate placeholder shapes before inserting dynamic content. | Extract a particular shape for export or metadata attachment.
// AI Prompts: Generate C# code that retrieves a shape by name with Aspose.Cells and handles a missing shape gracefully. | Show how to loop through all worksheet shapes and return the one matching a given name. | Explain how to rename a shape after obtaining it via the Shapes collection indexer.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle to a worksheet, assign a custom name, fetch the shape using the Shapes collection indexer, and confirm its existence with a null‑check before optionally saving the workbook.
class RetrieveShapeByName
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape and assign a name to it
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 0, 100, 100);
        shape.Name = "MyRectangle";

        // Retrieve the shape by its name
        Shape retrievedShape = worksheet.Shapes["MyRectangle"];

        // Verify that the shape exists
        if (retrievedShape != null)
        {
            Console.WriteLine("Shape found: " + retrievedShape.Name);
        }
        else
        {
            Console.WriteLine("Shape not found.");
        }

        // Save the workbook (optional)
        workbook.Save("RetrieveShapeDemo.xlsx");
    }
}

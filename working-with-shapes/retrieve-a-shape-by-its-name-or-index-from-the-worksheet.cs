// Title: Aspose.Cells for .NET – Retrieve a Worksheet Shape by Index or Name (C#)
// Description: This C# example creates a workbook, adds a rectangle and an oval, then shows how to obtain a shape from the worksheet either by its zero‑based index or by its default name (e.g., "Rectangle 1"). The sample also demonstrates changing the shape's fill colour and saving the file.
// Keywords: Aspose.Cells | C# | .NET | retrieve shape by index | retrieve shape by name | worksheet Shapes collection | Aspose.Cells shape example | modify shape fill color | Aspose.Cells API | RetrieveShapeDemo.cs | Excel shape manipulation | default shape name
// Common Searches: Aspose.Cells get shape by index C# | retrieve worksheet shape by name Aspose.Cells | change shape fill colour Aspose.Cells .NET | C# example using Shapes collection in Aspose.Cells | how to access "Rectangle 1" shape Aspose.Cells | Aspose.Cells Shapes[0] usage
// Developer Intent: Locate and manipulate a specific shape in a worksheet using its index or default name.
// Use Cases: Read the Type and Name of the first shape (index 0) for reporting. | Find a shape named "Rectangle 1" and set its FillFormat.ForeColor to green. | Validate that a shape lookup by name returns a non‑null object before applying changes. | Move a shape retrieved by index to a new cell range. | Iterate through the Shapes collection to apply bulk formatting.
// AI Prompts: Generate C# code that retrieves a shape by its name from an Aspose.Cells worksheet and changes its line style to dashed. | Show how to get a shape by index, reposition it to row 5 column 3, and save the workbook using Aspose.Cells. | Explain best practices for handling a missing shape name when accessing the Shapes collection in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a rectangle and an oval, then shows how to obtain a shape from the worksheet either by its zero‑based index or by its default name (e.g., "Rectangle 1"). The sample also demonstrates changing the shape's fill colour and saving the file.
class RetrieveShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some shapes to the worksheet
        // Rectangle with ID 1
        worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
        // Oval with ID 2
        worksheet.Shapes.AddOval(2, 0, 150, 100, 100, 100);

        // Retrieve a shape by its zero‑based index
        Shape shapeByIndex = worksheet.Shapes[0];
        Console.WriteLine($"Shape at index 0: Type={shapeByIndex.Type}, Name={shapeByIndex.Name}");

        // Retrieve a shape by its default name ("Rectangle 1", "Oval 2", etc.)
        Shape shapeByName = worksheet.Shapes["Rectangle 1"];
        if (shapeByName != null)
        {
            // Example modification: change fill color
            shapeByName.FillFormat.ForeColor = Color.Green;
            Console.WriteLine($"Shape retrieved by name: {shapeByName.Name}");
        }

        // Save the workbook to verify changes
        workbook.Save("RetrieveShapeDemo.xlsx");
    }
}

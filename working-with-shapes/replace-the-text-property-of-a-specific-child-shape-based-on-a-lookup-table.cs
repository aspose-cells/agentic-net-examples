// Title: Replace Shape Text Using a Lookup Dictionary in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load or create a workbook, iterate over all worksheet shapes, and substitute each shape's Text property with a new value from a Dictionary<string,string> before saving the file.
// Keywords: Aspose.Cells shape text replacement | C# replace worksheet shape caption | lookup dictionary Aspose.Cells | update Excel shape text .NET | iterate worksheet shapes Aspose
// Common Searches: How to change text of a shape in Aspose.Cells C# | Replace Excel shape captions using a dictionary | Iterate over worksheet shapes and modify Text property | Aspose.Cells update shape text based on lookup table
// Developer Intent: Swap the Text value of specific worksheet shapes according to entries in a lookup table.
// Use Cases: Refresh placeholder labels in a template workbook before distribution. | Localize shape captions by mapping original text to translated strings. | Synchronize shape text with external data sources during automated report generation.
// AI Prompts: Write C# code that loops through all shapes in an Aspose.Cells worksheet and replaces their Text using a Dictionary<string,string> lookup. | Show how to filter shapes that contain text and update only those whose text matches keys in a lookup table with Aspose.Cells for .NET. | Provide an example that saves the workbook after applying a lookup‑based text replacement to shape captions.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load or create a workbook, iterate over all worksheet shapes, and substitute each shape's Text property with a new value from a Dictionary<string,string> before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Sample shapes – in real scenario shapes already exist
        // -------------------------------------------------
        // Add a text box shape
        Shape textBox = worksheet.Shapes.AddTextBox(1, 0, 0, 100, 30, 200);
        textBox.Text = "Hello";

        // Add a rectangle shape
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 150, 30, 200);
        rectangle.Text = "World";

        // -------------------------------------------------
        // Lookup table: old text -> new text
        // -------------------------------------------------
        var lookup = new Dictionary<string, string>
        {
            { "Hello", "Hi" },
            { "World", "Earth" }
        };

        // -------------------------------------------------
        // Replace the Text property of each shape based on the lookup
        // -------------------------------------------------
        foreach (Shape shape in worksheet.Shapes)
        {
            // Ensure the shape actually contains text
            if (!string.IsNullOrEmpty(shape.Text) && lookup.ContainsKey(shape.Text))
            {
                // Replace with the corresponding value from the lookup table
                shape.Text = lookup[shape.Text];
            }
        }

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("ModifiedShapes.xlsx");
    }
}

// Title: Add a Named Rectangle Shape and Retrieve It Using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, inserts a rectangle shape, assigns a unique Name property, saves the file, then accesses the shape by its name to display its type and name.
// Keywords: Aspose.Cells C# | add shape Aspose.Cells | named shape Excel | retrieve shape by name | worksheet shape collection | .NET Excel automation | global Excel SDK
// Common Searches: Aspose.Cells add rectangle shape C# | How to set shape Name property in Aspose.Cells | Retrieve a shape by its Name in .NET Excel | Aspose.Cells shape naming example
// Developer Intent: Insert a shape, give it a unique identifier, and later reference it by that identifier.
// Use Cases: Mark a specific area in a report with a colored rectangle that can be updated later via its name. | Manage multiple interactive elements (buttons, icons) by assigning distinct names for individual manipulation. | Replace a placeholder shape with dynamic content such as a chart or image by locating it through its Name.
// AI Prompts: Write C# code that adds a circle shape, sets a custom Name, and later changes its fill color using Aspose.Cells. | Show how to iterate over all worksheet shapes, filter those whose Name starts with "Btn_", and output their types. | Provide an example that attaches a hyperlink to a shape identified by its Name property in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, inserts a rectangle shape, assigns a unique Name property, saves the file, then accesses the shape by its name to display its type and name.
class InsertShapeWithName
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: type, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // top row index
                10,   // vertical offset in pixels
                2,    // left column index
                10,   // horizontal offset in pixels
                100,  // height in pixels
                200); // width in pixels

            // Assign a unique name to the shape for later reference
            shape.Name = "MyUniqueShape";

            // Optionally set some visual properties
            shape.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
            // Line color setting removed due to API compatibility; you can customize other line properties if needed
            shape.Line.DashStyle = MsoLineDashStyle.Solid;
            shape.Line.Weight = 1.5;

            // Save the workbook (lifecycle save)
            string outputPath = "ShapeWithName.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");

            // Demonstrate retrieving the shape by its name
            Shape retrievedShape = worksheet.Shapes["MyUniqueShape"];
            Console.WriteLine("Retrieved shape type: " + retrievedShape.Type);
            Console.WriteLine("Shape name: " + retrievedShape.Name);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

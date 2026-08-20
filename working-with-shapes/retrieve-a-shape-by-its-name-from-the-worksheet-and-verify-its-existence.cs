// Title: Get a Worksheet Shape by Name and Verify Its Presence with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle named "MyRectangle" to the first worksheet, retrieves the shape using the ShapeCollection indexer (worksheet.Shapes["MyRectangle"]), checks for null, prints the shape's name and type, and saves the file.
// Keywords: Aspose.Cells | C# | shape by name | ShapeCollection indexer | retrieve worksheet shape | check shape existence | worksheet shapes API
// Common Searches: Aspose.Cells get shape by name C# | retrieve worksheet shape Aspose.Cells | check if shape exists Aspose.Cells .NET | ShapeCollection indexer example | find named shape in Excel using Aspose.Cells
// Developer Intent: Locate a shape on a worksheet using its assigned name and confirm that it is present before further processing.
// Use Cases: Modify properties of a specific named shape after it has been added. | Validate that required graphics are present before exporting or printing a workbook. | Attach additional formatting or data to a shape identified by its name.
// AI Prompts: Demonstrate safe retrieval of a worksheet shape by name in Aspose.Cells for .NET and handle the case when the shape is missing. | Show how to iterate through all shapes, find one with a given name, and then change its size or style using C#. | Explain whether the ShapeCollection indexer throws an exception for a non‑existent name and how to prevent errors.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeRetrievalDemo
{
    // Creates a workbook, adds a rectangle named "MyRectangle" to the first worksheet, retrieves the shape using the ShapeCollection indexer (worksheet.Shapes["MyRectangle"]), checks for null, prints the shape's name and type, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape and assign a custom name
            Shape addedShape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
            addedShape.Name = "MyRectangle";

            // Retrieve the shape by its name using the ShapeCollection indexer
            Shape retrievedShape = worksheet.Shapes["MyRectangle"];

            // Verify existence and output result
            if (retrievedShape != null)
            {
                Console.WriteLine($"Shape found: Name = {retrievedShape.Name}, Type = {retrievedShape.Type}");
            }
            else
            {
                Console.WriteLine("Shape with the specified name does not exist.");
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("ShapeRetrievalDemo.xlsx");
        }
    }
}

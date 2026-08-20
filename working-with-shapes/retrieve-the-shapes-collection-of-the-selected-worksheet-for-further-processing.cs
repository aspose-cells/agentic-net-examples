// Title: Aspose.Cells .NET – Retrieve Worksheet Shapes Collection and Add a Rectangle
// Description: Demonstrates how to access the Shapes collection of a worksheet in Aspose.Cells for .NET, enumerate existing shapes, add a new rectangle shape, and save the workbook.
// Keywords: Aspose.Cells Shapes .NET | worksheet.Shapes collection | ShapeCollection C# example | iterate worksheet shapes | add rectangle Aspose.Cells | Aspose.Cells shape manipulation | C# Aspose.Cells shape API
// Common Searches: How to get shapes from a worksheet using Aspose.Cells .NET | Aspose.Cells enumerate shapes C# | Add rectangle shape to Excel with Aspose.Cells | Aspose.Cells ShapeCollection example | C# code to list worksheet shapes Aspose
// Developer Intent: Access and work with the Shapes collection of a specific worksheet to read, modify, or add shapes.
// Use Cases: List all shapes on a worksheet and read their Type and Name. | Insert a rectangle shape at a defined row/column with custom dimensions. | Update a workbook after adding or modifying shapes.
// AI Prompts: Write C# code that obtains the ShapeCollection from a worksheet, loops through each shape, and prints its Type and Name using Aspose.Cells. | Show how to add a rectangle shape at row 2, column 2 with height 100 and width 200, then save the workbook with Aspose.Cells. | Explain how to filter shapes by Type after retrieving the ShapeCollection from a worksheet in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapesDemo
{
    // Demonstrates how to access the Shapes collection of a worksheet in Aspose.Cells for .NET, enumerate existing shapes, add a new rectangle shape, and save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook

            // Access the first worksheet (you can select any worksheet by index or name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the ShapeCollection from the selected worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Example: iterate through existing shapes (if any) for further processing
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                Console.WriteLine($"Shape {i}: Type={shape.Type}, Name={shape.Name}");
                // Additional processing can be done here
            }

            // (Optional) Add a shape to demonstrate that the collection is functional
            // shapes.AddRectangle(upperLeftRow, top, upperLeftColumn, left, height, width);
            // Example:
            shapes.AddRectangle(2, 0, 2, 0, 100, 200);

            // Save the workbook to verify changes (if needed)
            workbook.Save("ShapesDemo.xlsx");
        }
    }
}

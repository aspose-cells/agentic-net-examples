// Title: How to Retrieve Worksheet Shapes with Aspose.Cells for .NET (C#)
// Description: This example shows how to load or create a workbook, access a specific worksheet, obtain its ShapeCollection via the worksheet.Shapes property, loop through each shape to read its type and name, and finally save the file. The code works with Aspose.Cells for .NET in C#.
// Keywords: Aspose.Cells C# | worksheet shapes .NET | ShapeCollection Aspose | list Excel shapes | enumerate worksheet shapes | Aspose.Cells retrieve shapes | C# Excel shape iteration | Aspose.Cells API worksheet.Shapes | Excel shape properties C# | Aspose.Cells example
// Common Searches: Aspose.Cells get all shapes from a worksheet | C# retrieve ShapeCollection in Excel file | How to list shapes in an Excel sheet using Aspose | Iterate over worksheet shapes Aspose.Cells .NET | Access shape type and name with Aspose.Cells
// Developer Intent: Fetch the Shapes collection of a worksheet so it can be inspected, filtered, or modified programmatically.
// Use Cases: Log every shape’s type and name for audit purposes. | Locate and remove unwanted pictures or charts by name. | Adjust position, size, or formatting of shapes in bulk.
// AI Prompts: Write C# code that loads an existing workbook and deletes all picture shapes from the first worksheet using Aspose.Cells. | Provide an example that changes the fill color of every rectangle shape after retrieving the ShapeCollection. | Create a method that returns a dictionary of shape names and their corresponding types for a given worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example shows how to load or create a workbook, access a specific worksheet, obtain its ShapeCollection via the worksheet.Shapes property, loop through each shape to read its type and name, and finally save the file. The code works with Aspose.Cells for .NET in C#.
    class RetrieveShapesDemo
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new workbook; replace with new Workbook("input.xlsx") to load

            // Access the first worksheet (or any selected worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the Shapes collection from the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Example: iterate through all shapes and output basic info
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                Console.WriteLine($"Shape {i}: Type={shape.Type}, Name={shape.Name}");
            }

            // Save the workbook if any modifications were made
            workbook.Save("Output.xlsx");
        }
    }
}

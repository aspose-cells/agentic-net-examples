// Title: Group Shapes and Assign a Name with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle, an oval, and a line, groups them into a single GroupShape, sets the group's Name to "MyCompositeShape", and saves the file as GroupedShapesDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | GroupShape | group shapes | shape naming | composite shape | Aspose.Cells example | worksheet shapes
// Common Searches: Aspose.Cells group shapes C# | How to name a GroupShape in Aspose.Cells | Combine rectangle oval line into one shape Aspose.Cells | Assign custom name to grouped shapes Aspose.Cells .NET | Retrieve GroupShape by name Aspose.Cells
// Developer Intent: Combine multiple Shape objects into a single GroupShape and give the group a descriptive name for later reference.
// Use Cases: Design a reusable logo by grouping basic shapes and referencing it by name. | Create a diagram where all elements can be moved, hidden, or formatted together via the group name. | Locate and modify a previously named GroupShape across different worksheets or workbooks.
// AI Prompts: Generate C# code that adds a rectangle, an oval, and a line to a worksheet, groups them into a GroupShape, and sets the group's Name to "MyCompositeShape" using Aspose.Cells. | Show how to find a GroupShape by its Name in an existing workbook and change its position or size with Aspose.Cells for .NET. | Provide best‑practice error handling and resource cleanup when grouping shapes and assigning names in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle, an oval, and a line, groups them into a single GroupShape, sets the group's Name to "MyCompositeShape", and saves the file as GroupedShapesDemo.xlsx using Aspose.Cells for .NET.
    class GroupShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add three shapes to the worksheet
                Shape shape1 = sheet.Shapes.AddRectangle(2, 0, 2, 0, 60, 40);   // rectangle
                Shape shape2 = sheet.Shapes.AddOval(5, 0, 2, 0, 60, 40);        // oval
                Shape shape3 = sheet.Shapes.AddLine(8, 0, 2, 0, 100, 0);        // line

                // Group the three shapes into a composite object
                Shape[] shapesToGroup = new Shape[] { shape1, shape2, shape3 };
                GroupShape group = sheet.Shapes.Group(shapesToGroup);

                // Assign a descriptive name to the group for later reference
                group.Name = "MyCompositeShape";

                // Save the workbook
                workbook.Save("GroupedShapesDemo.xlsx");
                Console.WriteLine("Workbook saved successfully as GroupedShapesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            GroupShapesDemo.Run();
        }
    }
}

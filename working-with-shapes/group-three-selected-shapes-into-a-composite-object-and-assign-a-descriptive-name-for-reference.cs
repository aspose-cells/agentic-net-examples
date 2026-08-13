// Title: Group Multiple Shapes into a Named GroupShape with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle, an oval, and a line to the first worksheet, groups them into a single GroupShape, assigns the name "MyCompositeGroup", and saves the file as an XLSX document.
// Keywords: Aspose.Cells | GroupShape | shape grouping C# | named shape group | Aspose.Cells example | group multiple shapes | assign name to shape group | Excel workbook shapes
// Common Searches: how to group shapes in Aspose.Cells .NET | set name for GroupShape Aspose.Cells | C# example grouping rectangle oval line | retrieve GroupShape by name Aspose.Cells | save workbook after grouping shapes
// Developer Intent: Combine several worksheet shapes into a single GroupShape and give it a recognizable identifier for later manipulation.
// Use Cases: Build a flow‑chart component by grouping a rectangle, an oval, and a connector line, then move or resize the whole diagram as one object. | Apply a common fill or border style to all grouped elements with a single property change. | Locate the named GroupShape in a later operation to adjust its position, size, or to add additional shapes to the group.
// AI Prompts: Write C# code using Aspose.Cells that adds a rectangle, an oval, and a line, groups them into a GroupShape, sets the group's Name to "MyCompositeGroup", and saves the workbook. | Explain how to find a GroupShape by its Name property in an existing worksheet and modify its location or dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle, an oval, and a line to the first worksheet, groups them into a single GroupShape, assigns the name "MyCompositeGroup", and saves the file as an XLSX document.
    public class GroupThreeShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add three different shapes to the worksheet
                Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 60, 80);   // rectangle
                Shape oval = sheet.Shapes.AddOval(6, 0, 2, 0, 60, 80);       // oval
                Shape line = sheet.Shapes.AddLine(10, 0, 2, 0, 100, 0);      // line

                // Group the three shapes into a single composite object
                GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval, line });

                // Assign a descriptive name to the group for later reference
                group.Name = "MyCompositeGroup";

                // Save the workbook with the grouped shapes
                workbook.Save("GroupThreeShapesDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
